VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBlizzard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Module:       cQS5_tcm [cBlizzard.cls]
'  Title:        Unicode String Array Sorting Class
'  Salience:     Fast Case-Insensitive Sorting
'  Author:       Rohan Edwards // Rd
'  Email:        rde123•bigpond•com
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Description:
'   This string sorting class is extremely fast at case-insensitive
'   operations due to the combination of a highly optimized Blizzard
'   quicksort algorithm, extremely fast in-place unicode comparison
'   thanks to Ralph Eastwood, and a unique textual compare lookup
'   table that is built only once at initialization in less than
'   one millisecond thanks to LukeH, and stored in a long array
'   which is accessed in-line to minimize the case-conversion
'   impact on this supremely fast comparison algorithm.
'
'   Its performance on case-insensitive operations is only slightly
'   slower than its case-sensitive binary compare operations, and is
'   well over two times faster than StrComp's textual compare mode
'   when used within the very same algorithm.
'
'   Turning off strict compatibility with VB's StrComp function will
'   improve performance on case-insensitive operations even further.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Contributions:
'   QuickSort Origins:   C.A.R. Hoare
'   Blizzard Revision:   Rohan Edwards (Rd)
'   String Comparison:   Ralph Eastwood (tcm)
'                        McIlroy, Bostic and McIlroy
'   LCase Conversion:    LukeH (selftaught)
'   In-line Lookup:      Rohan Edwards (Rd)
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Features:
'   Capable of sorting arrays of millions of string items.
'   Capable of sorting in ascending and descending order.
'   Capable of case-sensitive and case-insensitive criteria.
'   Capable of any Lbound with equal or greater Ubound.
'   Capable of sorting sub-sets of the array data.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Requirements:
'   Array boundaries and data type limits are carefully managed
'   in this class. Therefore, advanced compiler optimizations
'   can be set however you see fit, with one exception - this
'   class will error if you set the Assume No Aliasing option.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Limitations:
'   Strings with embedded null characters will not compare beyond
'   the first occurence null character.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Licence:
'   You are free to use any part or all of this code even for
'   commercial purposes under the agreement that you receive
'   absolutely no warranties expressed or implied.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

' This class requires the following type library:

' Reference=*\G{C878CB53-7E75-4115-BD13-EECBC9430749}#1.0#0#MemAPIs.tlb#Memory APIs

' Or uncomment the following declares:

'Private Declare Sub CopyMemByV Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
'Private Declare Sub CopyMemByR Lib "kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)
'
'' SAFEARRAY Header, used in place of the real one to trick VB into letting us access string data in-place
'Private Type SAFEARRAY1D
'    cDims      As Integer      ' Count of dimensions in this array
'    fFeatures  As Integer      ' Bitfield flags indicating attributes of the array
'    cbElements As Long         ' Byte size of each element of the array
'    cLocks     As Long         ' Number of times the array has been locked without corresponding unlock. The cLocks field is a reference count that indicates how many times the array has been locked. When there is no lock, you're not supposed to access the array data, which is located in pvData.
'    pvData     As Long         ' Pointer to the start of the array data (use only if cLocks > 0)
'    cElements  As Long         ' Count of elements in this dimension
'    lLbound    As Long         ' The lower-bounding signed index number of this dimension
'    lUbound    As Long         ' The upper-bounding signed index number of this dimension
'End Type
'
'Private Const VT_BYREF = &H4000&        ' Tests whether the InitedArray routine was passed a Variant that contains an array, rather than directly an array in the former case ptr already points to the SA structure. Thanks to Monte Hansen for this fix
'Private Const FADF_AUTO = &H1           ' Array is allocated on the stack
'Private Const FADF_FIXEDSIZE = &H10     ' Array may not be resized or reallocated
'Private Const FADF_NO_REDIM = FADF_AUTO Or FADF_FIXEDSIZE
'
'' Used for unsigned arithmetic
'Private Const DW_MSB = &H80000000 ' DWord Most Significant Bit
'
'Public Enum eSortOrder
'    Descending = -1&
'    Ascending = 1&
'    #If False Then
'        Dim Descending, Ascending
'    #End If
'End Enum
'
'Public Enum eCompareMethod
'    BinaryCompare = 0&
'    TextCompare = 1&
'    #If False Then
'        Dim BinaryCompare, TextCompare
'    #End If
'End Enum
'
'Public Enum eCompare
'    Lesser = -1&
'    Equal = 0&
'    Greater = 1&
'    #If False Then
'        Dim Lesser, Equal, Greater
'    #End If
'End Enum

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Public Enum eOptimal
    ReversePretty = 1&     '0.001<<reverse-pretty-reverse>>0.002
    ReverseOptimal = 2&    '0.002<<reverse-sorting-unsorted>>0.003
    RefreshUnsorted = 3&   '0.003<<unsorted-refresh-sorting>>0.004
    #If False Then
        Dim ReversePretty, ReverseOptimal, RefreshUnsorted
    #End If
End Enum

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private StrPointers_Header As SAFEARRAY1D
Private StrPointers() As Long

Private CacheStr As SAFEARRAY1D
Private TempStr As SAFEARRAY1D
Private CacheChrs() As Integer
Private TempChrs() As Integer

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private StackLB() As Long       ' Pending boundary stacks
Private StackUB() As Long       ' Pending boundary stacks
Private StackSize As Long       ' Pending boundary stacks

Private TwisterBuf() As Long    ' Twister copymemory buffer
Private TwisterBufSize As Long  ' Twister copymemory buffer

Private ia() As Long

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Const Default_Order As Long = Ascending
Private Const Default_Method As Long = BinaryCompare

Private mMethod As eCompareMethod
Private mOrder As eSortOrder


'  The following properties should be set before sorting

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:  SortMethod [Public Property]
'  Purpose:    Determines the way strings are compared
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Property Get SortMethod() As eCompareMethod
    SortMethod = mMethod
End Property

Property Let SortMethod(ByVal eNewMethod As eCompareMethod)
    mMethod = eNewMethod
End Property

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:  SortOrder [Public Property]
'  Purpose:    Determines the order in which the strings are sorted
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Property Get SortOrder() As eSortOrder
    If mOrder = 0& Then mOrder = Default_Order
    SortOrder = mOrder
End Property

Property Let SortOrder(ByVal eNewOrder As eSortOrder)
    If eNewOrder = 0& Then
        If mOrder = 0& Then mOrder = Default_Order
    Else
        mOrder = eNewOrder
    End If
End Property

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:  BlizzardStringSort [Public Sub]
'
'  Features:
'   Non-recursive unicode string array quicksort algorithm using
'   a long array to access the string pointers in-place through
'   safearray header substitution.
'
'   In-line string comparison using integer arrays to access the
'   unicode string characters in-place through safearray header
'   substitution.
'
'   Case-insensitive textual comparison using a long array to
'   lookup unicode case-conversion in-line.
'
'  Contributions:
'   The Blizzard© quicksort algorithm created by Rohan Edwards.
'   String comparison by Ralph Eastwood/McIlroy, Bostic and McIlroy.
'   Fast unicode LCase conversion created by LukeH (selftaught).
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Sub BlizzardStringSort(sArray() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal StrCompCompatibility As Boolean = True)

    Dim I As Long

    If ubA - lbA < 1& Then Exit Sub

    With StrPointers_Header             ' Point to string array elements
        .pvData = VarPtr(sArray(lbA))
        .cElements = ubA - lbA + 1&
        .lLbound = lbA
    End With

    If SortOrder = Ascending Then
        Do While StrPointers(lbA) = 0&  ' Null string fix before main loops
            lbA = lbA + 1&
            If lbA > ubA Then Exit Do   ' Walk up while null strings
        Loop
    
        I = lbA + 1&
        Do Until I > ubA
            If StrPointers(I) = 0& Then ' Swap down any null strings
                StrPointers(I) = StrPointers(lbA)
                StrPointers(lbA) = 0&
                lbA = lbA + 1&          ' Set lb above null strings
            End If
            I = I + 1&
        Loop

    Else 'If SortOrder = Descending Then
        Do While StrPointers(ubA) = 0&  ' Null string fix before main loops
            ubA = ubA - 1&
            If ubA < lbA Then Exit Do   ' Walk down while null strings
        Loop
    
        I = ubA - 1&
        Do Until I < lbA
            If StrPointers(I) = 0& Then ' Swap up any null strings
                StrPointers(I) = StrPointers(ubA)
                StrPointers(ubA) = 0&
                ubA = ubA - 1&          ' Set ub below null strings
            End If
            I = I - 1&
        Loop
    End If

    If ubA - lbA < 1& Then Exit Sub     ' Sort from lb to ub

    Select Case SortOrder + SortMethod
        ' bin/dsc
        Case -1
            strSwap5bd lbA, ubA
        ' text/dsc
        Case 0
            strSwap5td lbA, ubA
            If StrCompCompatibility Then TwisterStringSort sArray, lbA, ubA, RefreshUnsorted
        ' bin/asc
        Case 1
            strSwap5ba lbA, ubA
        ' text/asc
        Case 2
            strSwap5ta lbA, ubA
            If StrCompCompatibility Then TwisterStringSort sArray, lbA, ubA, RefreshUnsorted
    End Select

End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure: TwisterStringSort [Public Sub]
'
'  Description:
'   Stable Insert/Binary hybrid using CopyMemory
'
'   Used internally as a finisher for case-insensitive operations
'   to remove differences between LCase and StrComp conversion.
'
'   Available as a public routine for refresh sorting operations.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Sub TwisterStringSort(sA() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal eOptimalA As eOptimal = RefreshUnsorted)
    Dim walk As Long
    Dim find As Long
    Dim midd As Long
    Dim base As Long
    Dim ceil As Long
    Dim item As String
    Dim cast As Long
    Dim mezz As Long
    Dim run As Long
    Dim lpStr As Long
    Dim lp As Long
    Dim lpB As Long
    Dim idx As Long
    Dim optimal As Long
    Dim eComp As eCompare
    
    walk = (ubA - lbA) + 1&                                       ' Grab array item count
    If walk < 2& Then Exit Sub                                    ' If nothing to do then exit
    eComp = SortOrder                                             ' Initialize compare variable
    optimal = InitializeBuffers(walk, eOptimalA)                  ' Initialize working buffers
    lpStr = VarPtr(item)                                          ' Cache pointer to the string variable
    lp = Sum(VarPtr(sA(lbA)), -(lbA * 4&))                        ' Cache pointer to the array
    lpB = VarPtr(TwisterBuf(0&))                                  ' Cache pointer to the buffer
    walk = lbA                                                    ' Initialize our walker variables
    mezz = ubA                                                    ' Initialize our walker variables
    Do Until walk = mezz ' ----==============================---- ' Do the twist while there's more items
        walk = walk + 1&                                          ' Walk up the array and use binary search to insert each item down into the sorted lower array
        CopyMemByV lpStr, Sum(lp, walk * 4&), 4&                  ' Grab current value into item
        find = walk                                               ' Default to current position
        ceil = walk - 1&                                          ' Set ceiling to current position - 1
        base = lbA                                                ' Set base to lower bound
        Do While StrComp(sA(ceil), item, mMethod) = eComp   '  .  ' While current item must move down
            midd = (base + ceil) \ 2&                             ' Find mid point
            Do Until StrComp(sA(midd), item, mMethod) = eComp     ' Step back up if equal or below
                base = midd + 1&                                  ' Bring up the base
                midd = (base + ceil) \ 2&                         ' Find mid point
                If midd = ceil Then Exit Do                       ' If we're up to ceiling
            Loop                                                  ' Out of loop > target pos
            find = midd                                           ' Set provisional to new ceiling
            If find = base Then Exit Do                           ' If we're down to base
            ceil = midd - 1&                                      ' Bring down the ceiling
        Loop '-Twister v4 ©Rd-     .      . ...  .             .  ' Out of stable binary search loops
        If find < walk Then                                       ' If current item needs to move down
            CopyMemByV lpStr, Sum(lp, find * 4&), 4&
            run = walk + 1&
            Do Until run > mezz Or run - walk > optimal           ' Runner do loop
                If Not (StrComp(item, sA(run), mMethod) = eComp) Then Exit Do
                run = run + 1&
            Loop
            cast = run - walk
            CopyMemByV lpB, Sum(lp, walk * 4&), cast * 4&         ' Grab current value(s)
            CopyMemByV Sum(lp, (find + cast) * 4&), Sum(lp, find * 4&), (walk - find) * 4& ' Move up items
            CopyMemByV Sum(lp, find * 4&), lpB, cast * 4&         ' Re-assign current value(s) into found pos
            If cast > 1& Then
                If Not run > mezz Then
                    idx = idx + 1&
                    StackLB(idx) = run - 1&                       ' Will increment back
                    StackUB(idx) = mezz
                End If
                walk = find
                mezz = find + cast - 1&
            End If
        End If
        If walk = mezz Then
            If idx Then
                walk = StackLB(idx)
                mezz = StackUB(idx)
                idx = idx - 1&
            End If
        End If
    Loop                                                          ' Out of walker do loop
    CopyMemByR ByVal lpStr, 0&, 4&                                ' De-reference pointer to item variable
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure: BlizzardStringSortIndexed [Public Sub]
'
'  Description:
'   Indexed non-recursive Blizzard based quicksort.
'
'   Receives a dynamic (redimmable) long array that holds references
'   to the string arrays indices. This is known as an indexed sort.
'   No changes are made to the source string array.
'
'   After a sort procedure is run the long array is ready as a sorted
'   index (lookup table) to the string array items.
'
'   E.G sA(idxA(lo)) returns the lo item in the string array whose
'   index may be anywhere in the string array.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Sub BlizzardStringSortIndexed(sArray() As String, iArray() As Long, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal StrCompCompatibility As Boolean = True)

    Dim I       As Long
    Dim cache   As Long

    If ubA - lbA < 1& Then Exit Sub
    
    ValidateIndexArray iArray, lbA, ubA  ' Validate the index array

    With StrPointers_Header          ' Point to string array elements
      .pvData = VarPtr(sArray(lbA))
      .cElements = ubA - lbA + 1&
      .lLbound = lbA
    End With

    If SortOrder = Ascending Then
        Do While StrPointers(iArray(lbA)) = 0& ' Null string fix before main loops
            lbA = lbA + 1&
            If lbA > ubA Then Exit Do  ' Walk up while null strings
        Loop
    
        I = lbA + 1&
        Do Until I > ubA
            If StrPointers(iArray(I)) = 0& Then
                cache = iArray(I)      ' Swap down any null strings
                iArray(I) = iArray(lbA)
                iArray(lbA) = cache
                lbA = lbA + 1&         ' Set lb above null strings
            End If
            I = I + 1&
        Loop

    Else 'If SortOrder = Descending Then
        Do While StrPointers(iArray(ubA)) = 0& ' Null string fix before main loops
            ubA = ubA - 1&
            If ubA < lbA Then Exit Do  ' Walk down while null strings
        Loop
    
        I = ubA - 1&
        Do Until I < lbA
            If StrPointers(iArray(I)) = 0& Then
                cache = iArray(I)      ' Swap up any null strings
                iArray(I) = iArray(ubA)
                iArray(ubA) = cache
                ubA = ubA - 1&         ' Set ub below null strings
            End If
            I = I - 1&
        Loop
    End If

    If ubA - lbA < 1& Then Exit Sub ' Sort from lb to ub

    Select Case SortOrder + SortMethod
      Case -1:  strSwap5bdi iArray, lbA, ubA ' bin/dsc
      Case 0:   strSwap5tdi iArray, lbA, ubA ' text/dsc
                If StrCompCompatibility Then TwisterStringSortIndexed sArray, iArray, lbA, ubA, RefreshUnsorted
      Case 1:   strSwap5bai iArray, lbA, ubA ' bin/asc
      Case 2:   strSwap5tai iArray, lbA, ubA ' text/asc
                If StrCompCompatibility Then TwisterStringSortIndexed sArray, iArray, lbA, ubA, RefreshUnsorted
    End Select

End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure: TwisterStringSortIndexed [Public Sub]
'
'  Description:
'   Indexed Stable Insert/Binary hybrid using CopyMemory
'
'   Receives a dynamic (redimmable) long array that holds references
'   to the string arrays indices. This is known as an indexed sort.
'   No changes are made to the source string array.
'
'   After a sort procedure is run the long array is ready as a sorted
'   index (lookup table) to the string array items.
'
'   E.G sA(idxA(lo)) returns the lo item in the string array whose
'   index may be anywhere in the string array.
'
'   Available as a public routine for refresh sorting operations.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Sub TwisterStringSortIndexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal eOptimalA As eOptimal = RefreshUnsorted)
    Dim walk As Long, find As Long, midd As Long
    Dim base As Long, ceil As Long, item As String
    Dim cast As Long, mezz As Long, run As Long
    Dim lpStr As Long, lpS As Long, lpB As Long
    Dim idx As Long, optimal As Long
    Dim eComp As eCompare, lpI As Long
    walk = (ubA - lbA) + 1&                      ' Grab array item count
    If walk < 2& Then Exit Sub                   ' If nothing to do then exit
    eComp = SortOrder                            ' Initialize compare variable
    ValidateIndexArray idxA, lbA, ubA            ' Validate the index array
    optimal = InitializeBuffers(walk, eOptimalA) ' Initialize working buffers
    lpStr = VarPtr(item)                                              ' Cache pointer to the string variable
    lpS = Sum(VarPtr(sA(lbA)), -(lbA * 4&))                           ' Cache pointer to the array
    lpI = Sum(VarPtr(idxA(lbA)), -(lbA * 4&))                         ' Cache pointer to the index array
    lpB = VarPtr(TwisterBuf(0&))                                      ' Cache pointer to the buffer
    walk = lbA: mezz = ubA                                            ' Initialize our walker variables
    Do Until walk = mezz ' ----==================================---- ' Do the twist while there's more items
        walk = walk + 1&                                              ' Walk up the array and use binary search to insert each item down into the sorted lower array
        CopyMemByV lpStr, Sum(lpS, idxA(walk) * 4&), 4&               ' Grab current value into item
        find = walk                                                   ' Default to current position
        ceil = walk - 1&                                              ' Set ceiling to current position - 1
        base = lbA                                                    ' Set base to lower bound
        Do While StrComp(sA(idxA(ceil)), item, mMethod) = eComp '  .  ' While current item must move down
            midd = (base + ceil) \ 2&                                 ' Find mid point
            Do Until StrComp(sA(idxA(midd)), item, mMethod) = eComp   ' Step back up if equal or below
                base = midd + 1&                                      ' Bring up the base
                midd = (base + ceil) \ 2&                             ' Find mid point
                If midd = ceil Then Exit Do                           ' If we're up to ceiling
            Loop                                                      ' Out of loop > target pos
            find = midd                                               ' Set provisional to new ceiling
            If find = base Then Exit Do                               ' If we're down to base
            ceil = midd - 1&                                          ' Bring down the ceiling
        Loop '-Twister v4i ©Rd-    .       . ...   .               .  ' Out of stable binary search loops
        If (find < walk) Then                                         ' If current item needs to move down
            CopyMemByV lpStr, Sum(lpS, idxA(find) * 4&), 4&
            run = walk + 1&
            Do Until run > mezz Or run - walk > optimal               ' Runner do loop
                If Not (StrComp(item, sA(idxA(run)), mMethod) = eComp) Then Exit Do
                run = run + 1&
            Loop: cast = (run - walk)
            CopyMemByV lpB, Sum(lpI, walk * 4&), cast * 4&            ' Grab current value(s)
            CopyMemByV Sum(lpI, (find + cast) * 4&), Sum(lpI, find * 4&), (walk - find) * 4& ' Move up items
            CopyMemByV Sum(lpI, find * 4&), lpB, cast * 4&            ' Re-assign current value(s) into found pos
            If cast > 1& Then
                If Not run > mezz Then
                    idx = idx + 1&
                    StackLB(idx) = run - 1&  ' Will increment back
                    StackUB(idx) = mezz
                End If
                walk = find
                mezz = find + cast - 1&
        End If: End If
        If walk = mezz Then
            If idx Then
                walk = StackLB(idx)
                mezz = StackUB(idx)
                idx = idx - 1&
    End If: End If: Loop           ' Out of walker do loop
    CopyMemByR ByVal lpStr, 0&, 4& ' De-reference pointer to item variable
    ' ----=================----
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Private non-recursive quicksort routines.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub strSwap5ba(ByVal lbA As Long, ByVal ubA As Long)
    Dim iChar As Long
    Dim cnt As Long
    Dim lo As Long
    Dim hi As Long
    
    On Error GoTo UnInit
    'On Error Resume Next
    
    InitializeStacks ubA - lbA                  ' Initialize pending boundary stacks
    Do
        hi = (ubA + lbA) \ 2&                   ' Get pivot index position
        CacheStr.pvData = StrPointers(hi)       ' Cache mid pivot
        StrPointers(hi) = StrPointers(ubA)      ' Set current
        lo = lbA
        hi = ubA                                ' Set bounds
        Do Until hi = lo                        ' Storm right in
            TempStr.pvData = StrPointers(lo)
            iChar = 0&
            Do While TempChrs(iChar) And (TempChrs(iChar) = CacheChrs(iChar))
               iChar = iChar + 1&
            Loop
            If Not CacheChrs(iChar) > TempChrs(iChar) Then ' (cache <= lo)
                StrPointers(hi) = StrPointers(lo)
                hi = hi - 1&
                Do Until hi = lo
                    TempStr.pvData = StrPointers(hi)
                    iChar = 0&
                    Do While TempChrs(iChar) And (TempChrs(iChar) = CacheChrs(iChar))
                        iChar = iChar + 1&
                    Loop
                    If Not TempChrs(iChar) > CacheChrs(iChar) Then ' (hi <= cache)
                        StrPointers(lo) = StrPointers(hi)
                        Exit Do
                    End If
                    hi = hi - 1&
                Loop
                If hi = lo Then Exit Do         ' Found swaps or out of loops
            End If
            lo = lo + 1&
        Loop '-Blizzard v5 ©Rd-
        
        StrPointers(hi) = CacheStr.pvData       ' Re-assign current
        If lbA < lo - 1& Then
            If ubA > lo + 1& Then
                cnt = cnt + 1&
                StackLB(cnt) = lo + 1&
                StackUB(cnt) = ubA
            End If
            ubA = lo - 1&
        ElseIf ubA > lo + 1& Then
            lbA = lo + 1&
        Else
            If cnt = 0& Then Exit Do
            lbA = StackLB(cnt)
            ubA = StackUB(cnt)
            cnt = cnt - 1&
        End If
    Loop
    
    Exit Sub
    
UnInit:
'Debug.Print "strSwap5ba: lbA=" & lbA & " ubA=" & ubA & " cnt=" & cnt & " lo=" & lo & " hi=" & hi & " iChar=" & iChar
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub strSwap5ta(ByVal lbA As Long, ByVal ubA As Long)
    Dim iChar As Long, cnt As Long
    Dim lo As Long, hi As Long
    InitializeStacks ubA - lbA              ' Initialize pending boundary stacks
    Do: hi = (ubA + lbA) \ 2&               ' Get pivot index position
        CacheStr.pvData = StrPointers(hi)   ' Cache mid pivot
        StrPointers(hi) = StrPointers(ubA)  ' Set current
        lo = lbA: hi = ubA                  ' Set bounds
        Do Until (hi = lo)                  ' Storm right in
            TempStr.pvData = StrPointers(lo)
            iChar = 0&
            Do While TempChrs(iChar) And (ia(TempChrs(iChar)) = ia(CacheChrs(iChar)))
               iChar = iChar + 1&
            Loop
            If Not ia(CacheChrs(iChar)) > ia(TempChrs(iChar)) Then ' (cache <= lo)
                StrPointers(hi) = StrPointers(lo)
                hi = hi - 1&
                Do Until (hi = lo)
                    TempStr.pvData = StrPointers(hi)
                    iChar = 0&
                    Do While TempChrs(iChar) And (ia(TempChrs(iChar)) = ia(CacheChrs(iChar)))
                        iChar = iChar + 1&
                    Loop
                    If Not ia(TempChrs(iChar)) > ia(CacheChrs(iChar)) Then ' (hi <= cache)
                        StrPointers(lo) = StrPointers(hi)
                        Exit Do
                    End If
                    hi = hi - 1&
                Loop
                If (hi = lo) Then Exit Do ' Found swaps or out of loops
            End If
            lo = lo + 1&
        Loop '-Blizzard v5 ©Rd-
        StrPointers(hi) = CacheStr.pvData ' Re-assign current
        If (lbA < lo - 1&) Then
            If (ubA > lo + 1&) Then cnt = cnt + 1&: StackLB(cnt) = lo + 1&: StackUB(cnt) = ubA
            ubA = lo - 1&
        ElseIf (ubA > lo + 1&) Then
            lbA = lo + 1&
        Else
            If cnt = 0& Then Exit Do
            lbA = StackLB(cnt): ubA = StackUB(cnt): cnt = cnt - 1&
        End If
    Loop
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub strSwap5bd(ByVal lbA As Long, ByVal ubA As Long)
    Dim iChar As Long, cnt As Long
    Dim lo As Long, hi As Long
    InitializeStacks ubA - lbA              ' Initialize pending boundary stacks
    Do: hi = (ubA + lbA) \ 2&               ' Get pivot index position
        CacheStr.pvData = StrPointers(hi)   ' Cache mid pivot
        StrPointers(hi) = StrPointers(ubA)  ' Set current
        lo = lbA: hi = ubA                  ' Set bounds
        Do Until (hi = lo)                  ' Storm right in
            TempStr.pvData = StrPointers(lo)
            iChar = 0&
            Do While TempChrs(iChar) And (TempChrs(iChar) = CacheChrs(iChar))
               iChar = iChar + 1&
            Loop
            If Not CacheChrs(iChar) < TempChrs(iChar) Then ' (cache >= lo)
                StrPointers(hi) = StrPointers(lo)
                hi = hi - 1&
                Do Until (hi = lo)
                    TempStr.pvData = StrPointers(hi)
                    iChar = 0&
                    Do While TempChrs(iChar) And (TempChrs(iChar) = CacheChrs(iChar))
                        iChar = iChar + 1&
                    Loop
                    If Not TempChrs(iChar) < CacheChrs(iChar) Then ' (hi >= cache)
                        StrPointers(lo) = StrPointers(hi)
                        Exit Do
                    End If
                    hi = hi - 1&
                Loop
                If (hi = lo) Then Exit Do ' Found swaps or out of loops
            End If
            lo = lo + 1&
        Loop '-Blizzard v5 ©Rd-
        StrPointers(hi) = CacheStr.pvData ' Re-assign current
        If (lbA < lo - 1&) Then
            If (ubA > lo + 1&) Then cnt = cnt + 1&: StackLB(cnt) = lo + 1&: StackUB(cnt) = ubA
            ubA = lo - 1&
        ElseIf (ubA > lo + 1&) Then
            lbA = lo + 1&
        Else
            If cnt = 0& Then Exit Do
            lbA = StackLB(cnt): ubA = StackUB(cnt): cnt = cnt - 1&
        End If
    Loop
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub strSwap5td(ByVal lbA As Long, ByVal ubA As Long)
    Dim iChar As Long, cnt As Long
    Dim lo As Long, hi As Long
    InitializeStacks ubA - lbA              ' Initialize pending boundary stacks
    Do: hi = (ubA + lbA) \ 2&               ' Get pivot index position
        CacheStr.pvData = StrPointers(hi)   ' Cache mid pivot
        StrPointers(hi) = StrPointers(ubA)  ' Set current
        lo = lbA: hi = ubA                  ' Set bounds
        Do Until (hi = lo)                  ' Storm right in
            TempStr.pvData = StrPointers(lo)
            iChar = 0&
            Do While TempChrs(iChar) And (ia(TempChrs(iChar)) = ia(CacheChrs(iChar)))
               iChar = iChar + 1&
            Loop
            If Not ia(CacheChrs(iChar)) < ia(TempChrs(iChar)) Then ' (cache >= lo)
                StrPointers(hi) = StrPointers(lo)
                hi = hi - 1&
                Do Until (hi = lo)
                    TempStr.pvData = StrPointers(hi)
                    iChar = 0&
                    Do While TempChrs(iChar) And (ia(TempChrs(iChar)) = ia(CacheChrs(iChar)))
                        iChar = iChar + 1&
                    Loop
                    If Not ia(TempChrs(iChar)) < ia(CacheChrs(iChar)) Then ' (hi >= cache)
                        StrPointers(lo) = StrPointers(hi)
                        Exit Do
                    End If
                    hi = hi - 1&
                Loop
                If (hi = lo) Then Exit Do ' Found swaps or out of loops
            End If
            lo = lo + 1&
        Loop '-Blizzard v5 ©Rd-
        StrPointers(hi) = CacheStr.pvData ' Re-assign current
        If (lbA < lo - 1&) Then
            If (ubA > lo + 1&) Then cnt = cnt + 1&: StackLB(cnt) = lo + 1&: StackUB(cnt) = ubA
            ubA = lo - 1&
        ElseIf (ubA > lo + 1&) Then
            lbA = lo + 1&
        Else
            If cnt = 0& Then Exit Do
            lbA = StackLB(cnt): ubA = StackUB(cnt): cnt = cnt - 1&
        End If
    Loop
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Private indexed non-recursive quicksort routines.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub strSwap5bai(idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    Dim iChar As Long, cnt As Long
    Dim lo As Long, hi As Long
    Dim idxItem As Long
    InitializeStacks ubA - lbA                   ' Initialize pending boundary stacks
    Do: hi = (ubA + lbA) \ 2&                    ' Get pivot index position
        idxItem = idxA(hi): idxA(hi) = idxA(ubA) ' Grab current index
        CacheStr.pvData = StrPointers(idxItem)   ' Cache mid pivot
        lo = lbA: hi = ubA                       ' Set bounds
        Do Until (hi = lo)                       ' Storm right in
            TempStr.pvData = StrPointers(idxA(lo))
            iChar = 0&
            Do While TempChrs(iChar) And (TempChrs(iChar) = CacheChrs(iChar))
               iChar = iChar + 1&
            Loop
            If Not CacheChrs(iChar) > TempChrs(iChar) Then ' (cache <= lo)
                idxA(hi) = idxA(lo)
                hi = hi - 1&
                Do Until (hi = lo)
                    TempStr.pvData = StrPointers(idxA(hi))
                    iChar = 0&
                    Do While TempChrs(iChar) And (TempChrs(iChar) = CacheChrs(iChar))
                        iChar = iChar + 1&
                    Loop
                    If Not TempChrs(iChar) > CacheChrs(iChar) Then ' (hi <= cache)
                        idxA(lo) = idxA(hi)
                        Exit Do
                    End If
                    hi = hi - 1&
                Loop
                If (hi = lo) Then Exit Do ' Found swaps or out of loops
            End If
            lo = lo + 1&
        Loop '-Blizzard v5i ©Rd-
        idxA(hi) = idxItem                ' Re-assign current
        If (lbA < lo - 1&) Then
            If (ubA > lo + 1&) Then cnt = cnt + 1&: StackLB(cnt) = lo + 1&: StackUB(cnt) = ubA
            ubA = lo - 1&
        ElseIf (ubA > lo + 1&) Then
            lbA = lo + 1&
        Else
            If cnt = 0& Then Exit Do
            lbA = StackLB(cnt): ubA = StackUB(cnt): cnt = cnt - 1&
        End If
    Loop
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub strSwap5tai(idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    Dim iChar As Long, cnt As Long
    Dim lo As Long, hi As Long
    Dim idxItem As Long
    InitializeStacks ubA - lbA                   ' Initialize pending boundary stacks
    Do: hi = (ubA + lbA) \ 2&                    ' Get pivot index position
        idxItem = idxA(hi): idxA(hi) = idxA(ubA) ' Grab current index
        CacheStr.pvData = StrPointers(idxItem)   ' Cache mid pivot
        lo = lbA: hi = ubA                       ' Set bounds
        Do Until (hi = lo)                       ' Storm right in
            TempStr.pvData = StrPointers(idxA(lo))
            iChar = 0&
            Do While TempChrs(iChar) And (ia(TempChrs(iChar)) = ia(CacheChrs(iChar)))
               iChar = iChar + 1&
            Loop
            If Not ia(CacheChrs(iChar)) > ia(TempChrs(iChar)) Then ' (cache <= lo)
                idxA(hi) = idxA(lo)
                hi = hi - 1&
                Do Until (hi = lo)
                    TempStr.pvData = StrPointers(idxA(hi))
                    iChar = 0&
                    Do While TempChrs(iChar) And (ia(TempChrs(iChar)) = ia(CacheChrs(iChar)))
                        iChar = iChar + 1&
                    Loop
                    If Not ia(TempChrs(iChar)) > ia(CacheChrs(iChar)) Then ' (hi <= cache)
                        idxA(lo) = idxA(hi)
                        Exit Do
                    End If
                    hi = hi - 1&
                Loop
                If (hi = lo) Then Exit Do ' Found swaps or out of loops
            End If
            lo = lo + 1&
        Loop '-Blizzard v5i ©Rd-
        idxA(hi) = idxItem                ' Re-assign current
        If (lbA < lo - 1&) Then
            If (ubA > lo + 1&) Then cnt = cnt + 1&: StackLB(cnt) = lo + 1&: StackUB(cnt) = ubA
            ubA = lo - 1&
        ElseIf (ubA > lo + 1&) Then
            lbA = lo + 1&
        Else
            If cnt = 0& Then Exit Do
            lbA = StackLB(cnt): ubA = StackUB(cnt): cnt = cnt - 1&
        End If
    Loop
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub strSwap5bdi(idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    Dim iChar As Long, cnt As Long
    Dim lo As Long, hi As Long
    Dim idxItem As Long
    InitializeStacks ubA - lbA                   ' Initialize pending boundary stacks
    Do: hi = (ubA + lbA) \ 2&                    ' Get pivot index position
        idxItem = idxA(hi): idxA(hi) = idxA(ubA) ' Grab current index
        CacheStr.pvData = StrPointers(idxItem)   ' Cache mid pivot
        lo = lbA: hi = ubA                       ' Set bounds
        Do Until (hi = lo)                       ' Storm right in
            TempStr.pvData = StrPointers(idxA(lo))
            iChar = 0&
            Do While TempChrs(iChar) And (TempChrs(iChar) = CacheChrs(iChar))
               iChar = iChar + 1&
            Loop
            If Not CacheChrs(iChar) < TempChrs(iChar) Then ' (cache >= lo)
                idxA(hi) = idxA(lo)
                hi = hi - 1&
                Do Until (hi = lo)
                    TempStr.pvData = StrPointers(idxA(hi))
                    iChar = 0&
                    Do While TempChrs(iChar) And (TempChrs(iChar) = CacheChrs(iChar))
                        iChar = iChar + 1&
                    Loop
                    If Not TempChrs(iChar) < CacheChrs(iChar) Then ' (hi >= cache)
                        idxA(lo) = idxA(hi)
                        Exit Do
                    End If
                    hi = hi - 1&
                Loop
                If (hi = lo) Then Exit Do ' Found swaps or out of loops
            End If
            lo = lo + 1&
        Loop '-Blizzard v5i ©Rd-
        idxA(hi) = idxItem                ' Re-assign current
        If (lbA < lo - 1&) Then
            If (ubA > lo + 1&) Then cnt = cnt + 1&: StackLB(cnt) = lo + 1&: StackUB(cnt) = ubA
            ubA = lo - 1&
        ElseIf (ubA > lo + 1&) Then
            lbA = lo + 1&
        Else
            If cnt = 0& Then Exit Do
            lbA = StackLB(cnt): ubA = StackUB(cnt): cnt = cnt - 1&
        End If
    Loop
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub strSwap5tdi(idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    Dim iChar As Long, cnt As Long
    Dim lo As Long, hi As Long
    Dim idxItem As Long
    InitializeStacks ubA - lbA                   ' Initialize pending boundary stacks
    Do: hi = (ubA + lbA) \ 2&                    ' Get pivot index position
        idxItem = idxA(hi): idxA(hi) = idxA(ubA) ' Grab current index
        CacheStr.pvData = StrPointers(idxItem)   ' Cache mid pivot
        lo = lbA: hi = ubA                       ' Set bounds
        Do Until (hi = lo)                       ' Storm right in
            TempStr.pvData = StrPointers(idxA(lo))
            iChar = 0&
            Do While TempChrs(iChar) And (ia(TempChrs(iChar)) = ia(CacheChrs(iChar)))
               iChar = iChar + 1&
            Loop
            If Not ia(CacheChrs(iChar)) < ia(TempChrs(iChar)) Then ' (cache >= lo)
                idxA(hi) = idxA(lo)
                hi = hi - 1&
                Do Until (hi = lo)
                    TempStr.pvData = StrPointers(idxA(hi))
                    iChar = 0&
                    Do While TempChrs(iChar) And (ia(TempChrs(iChar)) = ia(CacheChrs(iChar)))
                        iChar = iChar + 1&
                    Loop
                    If Not ia(TempChrs(iChar)) < ia(CacheChrs(iChar)) Then ' (hi >= cache)
                        idxA(lo) = idxA(hi)
                        Exit Do
                    End If
                    hi = hi - 1&
                Loop
                If (hi = lo) Then Exit Do ' Found swaps or out of loops
            End If
            lo = lo + 1&
        Loop '-Blizzard v5i ©Rd-
        idxA(hi) = idxItem                ' Re-assign current
        If (lbA < lo - 1&) Then
            If (ubA > lo + 1&) Then cnt = cnt + 1&: StackLB(cnt) = lo + 1&: StackUB(cnt) = ubA
            ubA = lo - 1&
        ElseIf (ubA > lo + 1&) Then
            lbA = lo + 1&
        Else
            If cnt = 0& Then Exit Do
            lbA = StackLB(cnt): ubA = StackUB(cnt): cnt = cnt - 1&
        End If
    Loop
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:   Class_Initialize [Private Sub]
'  Purpose:     Initialize the SafeArray Headers
'               Build the Unicode Conversion Lookup Table
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub Class_Initialize()
    mOrder = Default_Order
    mMethod = Default_Method

    With StrPointers_Header
        .cDims = 1                  ' 1 Dimensional
        .fFeatures = FADF_NO_REDIM  ' Cannot REDIM the array
        .cbElements = 4&            ' This is a long array
        .cLocks = 1&                ' Lock the array
    End With
    CopyMemByR ByVal ArrayPtr(StrPointers), VarPtr(StrPointers_Header), 4&

    With CacheStr
        .cDims = 1                  ' 1 Dimensional
        .fFeatures = FADF_NO_REDIM  ' Cannot REDIM the array
        .cbElements = 2&            ' This is an integer array
        .cElements = &H7FFFFFFF     ' Set max string length
        .cLocks = 1&                ' Lock the array
    End With
    CopyMemByR ByVal ArrayPtr(CacheChrs), VarPtr(CacheStr), 4&

    With TempStr
        .cDims = 1                  ' 1 Dimensional
        .fFeatures = FADF_NO_REDIM  ' Cannot REDIM the array
        .cbElements = 2&            ' This is an integer array
        .cElements = &H7FFFFFFF     ' Set max string length
        .cLocks = 1&                ' Lock the array
    End With
    CopyMemByR ByVal ArrayPtr(TempChrs), VarPtr(TempStr), 4&

    BuildUnicodeConversionLookupTable

End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:   Class_Terminate [Private Sub]
'  Purpose:     Nullify the SafeArray Headers
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub Class_Terminate()

    'Clear up the array pointers so vb doesn't try to free memory twice
    CopyMemByR ByVal ArrayPtr(TempChrs), 0&, 4&
    CopyMemByR ByVal ArrayPtr(CacheChrs), 0&, 4&
    CopyMemByR ByVal ArrayPtr(StrPointers), 0&, 4&

End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure: ArrayPtr [Public Function]
'
'  Description:
'   This function returns a pointer to the SAFEARRAY header of
'   any Visual Basic array, including a Visual Basic string array.
'
'   Substitutes both ArrPtr and StrArrPtr.
'
'   This function will work with vb5 or vb6 without modification.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Function ArrayPtr(Arr) As Long
    Dim iDataType As Integer
    On Error GoTo UnInit
    CopyMemByR iDataType, Arr, 2&                       ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then           ' if a valid array was passed
        CopyMemByR ArrayPtr, ByVal VarPtr(Arr) + 8&, 4& ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array. Thanks to Francesco Balena.
    End If
UnInit:
End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure: InitedArray [Public Function]
'
'  Description:
'   This function determines if the passed array is initialized,
'   and if so will return -1.
'
'   It will also optionally indicate whether the array can be redimmed;
'   in which case it will return -2.
'
'   If the array is uninitialized (never redimmed or has been erased)
'   it will return 0 (zero).
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Function InitedArray(Arr, lbA As Long, ubA As Long, Optional ByVal bTestRedimable As Boolean) As Long
    ' Thanks to Francesco Balena who solved the Variant headache,
    ' and to Monte Hansen for the ByRef fix
    Dim tSA As SAFEARRAY1D, lpSA As Long
    Dim lOffset As Long, iDataType As Integer
    On Error GoTo UnInit
    CopyMemByR iDataType, Arr, 2&                       ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then           ' if a valid array was passed
        CopyMemByR lpSA, ByVal Sum(VarPtr(Arr), 8&), 4& ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array
        If (iDataType And VT_BYREF) Then                ' see whether the function was passed a Variant that contains an array, rather than directly an array in the former case lpSA already points to the SA structure. Thanks to Monte Hansen for this fix
            CopyMemByR lpSA, ByVal lpSA, 4&             ' lpSA is a discripter (pointer) to the safearray structure
        End If
        InitedArray = (lpSA <> 0&)
        If InitedArray Then
            CopyMemByR tSA.cDims, ByVal lpSA, 4&
            If bTestRedimable Then ' Return -2 if redimmable
                InitedArray = InitedArray + ((tSA.fFeatures And FADF_FIXEDSIZE) <> FADF_FIXEDSIZE)
            End If '-©Rd-
            lOffset = 16& + ((tSA.cDims - 1&) * 8&)
            CopyMemByR tSA.cElements, ByVal Sum(lpSA, lOffset), 8&
            tSA.lUbound = tSA.lLbound + tSA.cElements - 1&
            If (lbA < tSA.lLbound) Then lbA = tSA.lLbound
            If (ubA > tSA.lUbound) Then ubA = tSA.lUbound
    End If: End If
UnInit:
End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure: ValidateIndexArray [Private Function]
'
'  Description:
'   Receives a dynamic (redimmable) long array to holds references
'   to the string arrays indices.
'
'   Ensures that the index array is in the required pre-sort state.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub ValidateIndexArray(idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    Dim bReDim As Boolean, bReDimPres As Boolean
    Dim LB As Long, ub As Long, J As Long
    LB = &H80000000: ub = &H7FFFFFFF
    bReDim = Not InitedArray(idxA, LB, ub)
    If bReDim = False Then
        bReDim = (lbA < LB)
        bReDimPres = (ubA > ub)
    End If
    If bReDim Then
        ReDim idxA(lbA To ubA) As Long
    ElseIf bReDimPres Then
        ReDim Preserve idxA(LB To ubA) As Long
    End If '-©Rd-
    If (idxA(ubA) = 0&) Then
        If (idxA(lbA) = 0&) Then
            For J = lbA To ubA
                idxA(J) = J
            Next
        ElseIf bReDimPres Then
            For J = ub + 1& To ubA
                idxA(J) = J
            Next
        End If
    End If
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure: Sum [Private Function]
'
'  Description:
'   Enables valid addition and subtraction of unsigned long ints.
'   Treats lPtr as an unsigned long and returns an unsigned long.
'   Allows safe arithmetic operations on memory address pointers.
'   Assumes valid pointer and pointer offset.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Function Sum(ByVal lPtr As Long, ByVal lOffset As Long) As Long
    If lOffset > 0& Then
        If lPtr And DW_MSB Then ' if ptr < 0
           Sum = lPtr + lOffset ' ignors > unsigned int max
        ElseIf (lPtr Or DW_MSB) < -lOffset Then
           Sum = lPtr + lOffset ' result is below signed int max
        Else                    ' result wraps to min signed int
           Sum = (lPtr + DW_MSB) + (lOffset + DW_MSB)
        End If
    ElseIf lOffset = 0& Then
        Sum = lPtr
    Else 'If lOffset < 0 Then
        If (lPtr And DW_MSB) = 0& Then ' if ptr > 0
           Sum = lPtr + lOffset ' ignors unsigned int < zero
        ElseIf (lPtr - DW_MSB) >= -lOffset Then
           Sum = lPtr + lOffset ' result is above signed int min
        Else                    ' result wraps to max signed int
           Sum = (lOffset - DW_MSB) + (lPtr - DW_MSB)
        End If
    End If
End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:   InitializeStacks [Private Sub]
'  Purpose:     Initialize the array boundary stacks
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub InitializeStacks(ByVal lSize As Long)
    Dim I As Long
    Dim J As Long
    
    I = 1&
    J = 32&                   ' Allow for likely stack count + 32
    Do Until I > lSize        ' 10000 -> 14 + 32
        J = J + 1&            ' 100000 -> 17 + 32
        I = I + I             ' 1000000 -> 20 + 32
    Loop                      ' 10000000 -> 24 + 32
    If J > StackSize Then
        StackSize = J
        ReDim StackLB(0& To J) As Long  ' Stack to hold pending lower boundries
        ReDim StackUB(0& To J) As Long  ' Stack to hold pending upper boundries
    End If
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:   InitializeBuffers [Private Function]
'  Purpose:     Calculates the range of the Twister's runners
'               Initializes pending runner stacks
'               Initializes the Twisters runner buffer
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Function InitializeBuffers(ByVal lCount As Long, ByVal eOpt As eOptimal) As Long
    Dim curve As Long, optimal As Long                       ' CraZy performance curve
    Const n10K As Long = 10000&                              ' .
    Const n20K As Long = 20000&                              '    .
    If lCount > n20K Then curve = 12& * (lCount \ n10K - 2&) '      .
    optimal = (lCount * (eOpt * 0.0012!)) - curve + 4&       '       .
    If optimal > StackSize Then
        StackSize = optimal
        ReDim StackLB(0& To optimal) As Long   ' Stack to hold pending lower boundries
        ReDim StackUB(0& To optimal) As Long   ' Stack to hold pending upper boundries
    End If
    If lCount > TwisterBufSize Then
        TwisterBufSize = lCount
        ReDim TwisterBuf(0& To lCount) As Long ' This is a cache used when moving ranges
    End If
    InitializeBuffers = optimal
End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:   BuildUnicodeConversionLookupTable [Private Sub]
'
'  Purpose:     Initializes the unicode lookup table to enable
'               in-line case-insensitive textual comparison.
'
'    CharLowerBuffW ia(lb), iLen    '0.0009881144111891 secs
'    ia(i) = LukeLCaseI             '0.0009657652020019 secs
'    ia(i) = AscW(LCase$(ChrW$(i))) '0.0092841408614782 secs
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub BuildUnicodeConversionLookupTable()
    Dim I As Long, J As Long
    ReDim ia(-32768 To 32767) As Long '65535
    For I = -32768 To -224
       ia(I) = I
    Next
    For I = -223 To 9423
      'ia(i) = AscW(LCase$(ChrW$(i)))
       J = I
       LukeLCaseI J          'For i = 32768 To 65535
       ia(I) = J             '   j = i               'For i = 32768 To 65535
    Next                     '   LukeLCaseI j        '   ia(i) = ia(i - 65536)
    For I = 9424 To 32767    '   ia(i) = j           'Next
       ia(I) = I             'Next
    Next
    ' StrComp compatibility ("_", "@", etc, list before number "0")
    For I = 15 To 31
        ia(I) = 14&     '14 (ctrl chars)
    Next
    For I = 32 To 47
        ia(I) = I - 17& '15...30
    Next
    For I = 58 To 64
        ia(I) = I - 27& '31...37
    Next
    For I = 91 To 96
        ia(I) = I - 53& '38...43
    Next
    For I = 123 To 126
        ia(I) = I - 79& '44...47
    Next
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:     LukeLCaseI [Private Sub]
'  Purpose:       Unicode LCase
'
'  Description:
'   Sub to enable case insensitive comparison of unicode strings
'   as integer data. Substitutes AscW(LCase$(ChrW$(i))) without
'   any type conversion. With native code, it is extremely fast.
'
'  By: Adapted directly from CharLower by selftaught
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Private Sub LukeLCaseI(ByRef piChar As Long)

    Const MakeLCase As Long = 32

    Const Neg223 As Long = -223
    Const Neg197 As Long = -197

    Const Zero As Long = 0
    Const One As Long = 1
    Const Two As Long = 2
    Const Eight As Long = 8

    Const n16 As Long = 16
    Const n26 As Long = 26
    Const n37 As Long = 37
    Const n48 As Long = 48
    Const n63 As Long = 63
    Const n65 As Long = 65
    Const n74 As Long = 74
    Const n80 As Long = 80
    Const n86 As Long = 86
    Const n91 As Long = 91
    Const n100 As Long = 100
    Const n112 As Long = 112
    Const n192 As Long = 192
    Const n205 As Long = 205
    Const n215 As Long = 215
    Const n216 As Long = 216
    Const n217 As Long = 217
    Const n218 As Long = 218
    Const n223 As Long = 223
    Const n126 As Long = 126
    Const n128 As Long = 128
    Const n255 As Long = 255
    Const n256 As Long = 256
    Const n304 As Long = 304
    Const n311 As Long = 311
    Const n313 As Long = 313
    Const n328 As Long = 328
    Const n330 As Long = 330
    Const n375 As Long = 375
    Const n376 As Long = 376
    Const n377 As Long = 377
    Const n379 As Long = 379
    Const n381 As Long = 381
    Const n385 As Long = 385
    Const n386 As Long = 386
    Const n388 As Long = 388
    Const n390 As Long = 390
    Const n391 As Long = 391
    Const n392 As Long = 392
    Const n393 As Long = 393
    Const n394 As Long = 394
    Const n395 As Long = 395
    Const n396 As Long = 396
    Const n398 As Long = 398
    Const n399 As Long = 399
    Const n400 As Long = 400
    Const n401 As Long = 401
    Const n402 As Long = 402
    Const n403 As Long = 403
    Const n404 As Long = 404
    Const n406 As Long = 406
    Const n407 As Long = 407
    Const n408 As Long = 408
    Const n409 As Long = 409
    Const n412 As Long = 412
    Const n413 As Long = 413
    Const n415 As Long = 415
    Const n416 As Long = 416
    Const n418 As Long = 418
    Const n420 As Long = 420
    Const n423 As Long = 423
    Const n425 As Long = 425
    Const n428 As Long = 428
    Const n430 As Long = 430
    Const n431 As Long = 431
    Const n433 As Long = 433
    Const n434 As Long = 434
    Const n435 As Long = 435
    Const n437 As Long = 437
    Const n439 As Long = 439
    Const n440 As Long = 440
    Const n444 As Long = 444
    Const n452 As Long = 452
    Const n455 As Long = 455
    Const n458 As Long = 458
    Const n461 As Long = 461
    Const n476 As Long = 476
    Const n477 As Long = 477
    Const n495 As Long = 495
    Const n497 As Long = 497
    Const n499 As Long = 499
    Const n500 As Long = 500
    Const n501 As Long = 501
    Const n506 As Long = 506
    Const n535 As Long = 535
    Const n595 As Long = 595
    Const n596 As Long = 596
    Const n601 As Long = 601
    Const n603 As Long = 603
    Const n608 As Long = 608
    Const n611 As Long = 611
    Const n616 As Long = 616
    Const n617 As Long = 617
    Const n623 As Long = 623
    Const n626 As Long = 626
    Const n629 As Long = 629
    Const n658 As Long = 658
    Const n902 As Long = 902
    Const n904 As Long = 904
    Const n907 As Long = 907
    Const n908 As Long = 908
    Const n910 As Long = 910
    Const n911 As Long = 911
    Const n913 As Long = 913
    Const n930 As Long = 930
    Const n940 As Long = 940
    Const n972 As Long = 972
    Const n994 As Long = 994
    Const n1007 As Long = 1007
    Const n1025 As Long = 1025
    Const n1037 As Long = 1037
    Const n1040 As Long = 1040
    Const n1072 As Long = 1072
    Const n1120 As Long = 1120
    Const n1153 As Long = 1153
    Const n1168 As Long = 1168
    Const n1215 As Long = 1215
    Const n1217 As Long = 1217
    Const n1219 As Long = 1219
    Const n1223 As Long = 1223
    Const n1227 As Long = 1227
    Const n1232 As Long = 1232
    Const n1259 As Long = 1259
    Const n1262 As Long = 1262
    Const n1269 As Long = 1269
    Const n1272 As Long = 1272
    Const n1273 As Long = 1273
    Const n1329 As Long = 1329
    Const n1366 As Long = 1366
    Const n4256 As Long = 4256
    Const n4294 As Long = 4294
    Const n7680 As Long = 7680
    Const n7829 As Long = 7829
    Const n7840 As Long = 7840
    Const n7929 As Long = 7929
    Const n7944 As Long = 7944
    Const n7952 As Long = 7952
    Const n7960 As Long = 7960
    Const n7966 As Long = 7966
    Const n7976 As Long = 7976
    Const n7984 As Long = 7984
    Const n7992 As Long = 7992
    Const n8000 As Long = 8000
    Const n8008 As Long = 8008
    Const n8014 As Long = 8014
    Const n8025 As Long = 8025
    Const n8032 As Long = 8032
    Const n8040 As Long = 8040
    Const n8048 As Long = 8048
    Const n8120 As Long = 8120
    Const n8121 As Long = 8121
    Const n8122 As Long = 8122
    Const n8123 As Long = 8123
    Const n8136 As Long = 8136
    Const n8140 As Long = 8140
    Const n8152 As Long = 8152
    Const n8153 As Long = 8153
    Const n8154 As Long = 8154
    Const n8155 As Long = 8155
    Const n8165 As Long = 8165
    Const n8168 As Long = 8168
    Const n8169 As Long = 8169
    Const n8170 As Long = 8170
    Const n8171 As Long = 8171
    Const n8172 As Long = 8172
    Const n8184 As Long = 8184
    Const n8185 As Long = 8185
    Const n8186 As Long = 8186
    Const n8187 As Long = 8187
    Const n8544 As Long = 8544
    Const n8560 As Long = 8560
    Const n9397 As Long = 9397
    Const n9423 As Long = 9423
    Const n32768 As Long = 32768
    Const n65536 As Long = 65536

    If piChar < Neg223 Then ' -32768 .. -222
    ElseIf piChar > n9423 Then ' 9424..(65535)
        If piChar < n32768 Then ' 9424..32767
        Else 'piChar > n32767  ' 32768..65535
            Dim I As Long
            I = piChar - n65536
            If I >= Neg223 And I <= n9423 Then
                LukeLCaseI I
            End If
            piChar = I
        End If
    ElseIf piChar < n4294 Then '..4293
        If piChar > n1366 And piChar < n4256 Then ' 1367..4255
        ElseIf piChar < n416 Then '..415
            If piChar < n192 Then '..191
                If piChar < Neg197 Then
                    piChar = piChar + MakeLCase
                ElseIf piChar < n65 Then
                ElseIf piChar < n91 Then
                    piChar = piChar + MakeLCase
                End If
            ElseIf piChar < n215 Then
                piChar = piChar + MakeLCase
            ElseIf piChar < n216 Then
            ElseIf piChar < n223 Then ' 216..222
                piChar = piChar + MakeLCase
            ElseIf piChar < n256 Then
            ElseIf piChar < n304 Then ' 256..303
                If piChar Mod Two = Zero Then piChar = piChar + One
            ElseIf piChar = n304 Then
            ElseIf piChar < n311 Then ' 305..310
                If piChar Mod Two = Zero Then piChar = piChar + One
            ElseIf piChar < n313 Then
            ElseIf piChar < n328 Then ' 313..327
                If piChar Mod Two Then piChar = piChar + One
            ElseIf piChar < n330 Then
            ElseIf piChar < n375 Then ' 330..374
                If piChar Mod Two = Zero Then piChar = piChar + One
            ElseIf piChar = n375 Then
            ElseIf piChar = n376 Then
                piChar = n255
            ElseIf piChar = n377 Or piChar = n379 Or piChar = n381 Then
                piChar = piChar + One
            ElseIf piChar < n385 Then
            ElseIf piChar = n385 Then
                piChar = n595
            ElseIf piChar = n386 Or piChar = n388 Then
                piChar = piChar + One
            ElseIf piChar = n390 Then
                piChar = n596
            ElseIf piChar = n391 Then
                piChar = n392
            ElseIf piChar = n393 Or piChar = n394 Then
                piChar = piChar + n205
            ElseIf piChar = n395 Then
                piChar = n396
            ElseIf piChar < n398 Then
            ElseIf piChar = n398 Then
                piChar = n477
            ElseIf piChar = n399 Then
                piChar = n601
            ElseIf piChar = n400 Then
                piChar = n603
            ElseIf piChar = n401 Then
                piChar = n402
            ElseIf piChar = n403 Then
                piChar = n608
            ElseIf piChar = n404 Then
                piChar = n611
            ElseIf piChar = n406 Then
                piChar = n617
            ElseIf piChar = n407 Then
                piChar = n616
            ElseIf piChar = n408 Then
                piChar = n409
            ElseIf piChar < n412 Then
            ElseIf piChar = n412 Then
                piChar = n623
            ElseIf piChar = n413 Then
                piChar = n626
            ElseIf piChar = n415 Then
                piChar = n629
            End If
        ElseIf piChar < n902 Then
            If piChar < n535 Then '..534
                If piChar = n416 Or piChar = n418 Or piChar = n420 Or piChar = n423 Then
                    piChar = piChar + One
                ElseIf piChar < n425 Then
                ElseIf piChar = n425 Or piChar = n430 Then
                    piChar = piChar + n218
                ElseIf piChar = n428 Or piChar = n431 Or piChar = n435 Or piChar = n437 Then
                    piChar = piChar + One
                ElseIf piChar = n433 Or piChar = n434 Then
                    piChar = piChar + n217
                ElseIf piChar = n439 Then
                    piChar = n658
                ElseIf piChar = n440 Or piChar = n444 Then
                    piChar = piChar + One
                ElseIf piChar < n452 Then
                ElseIf piChar = n452 Or piChar = n455 Or piChar = n458 Then
                    piChar = piChar + Two
                ElseIf piChar < n461 Then
                ElseIf piChar < n476 Then ' 461..475
                    If piChar Mod Two Then piChar = piChar + One
                ElseIf piChar <= n477 Then
                ElseIf piChar < n495 Then ' 478..494
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar = n497 Then
                    piChar = n499
                ElseIf piChar = n500 Then
                    piChar = n501
                ElseIf piChar < n506 Then
                Else 'If piChar < n535 Then ' 506..534
                    If piChar Mod Two = Zero Then piChar = piChar + One
                End If
            End If
        ElseIf piChar < n1329 Then '..1328
            If piChar < n1273 Then '..1272
                If piChar = n902 Then
                    piChar = n940
                ElseIf piChar < n904 Then
                ElseIf piChar < n907 Then ' 904..906
                    piChar = piChar + n37
                ElseIf piChar = n908 Then
                    piChar = n972
                ElseIf piChar = n910 Or piChar = n911 Then
                    piChar = piChar + n63
                ElseIf piChar < n913 Then
                ElseIf piChar = n930 Then
                ElseIf piChar < n940 Then ' 913..939
                    piChar = piChar + MakeLCase
                ElseIf piChar < n994 Then
                ElseIf piChar < n1007 Then ' 994..1006
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar < n1025 Then
                ElseIf piChar = n1037 Then
                ElseIf piChar < n1040 Then ' 1025..1039
                    piChar = piChar + n80
                ElseIf piChar < n1072 Then ' 1040..1071
                    piChar = piChar + MakeLCase
                ElseIf piChar < n1120 Then
                ElseIf piChar < n1153 Then ' 1120..1152
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar < n1168 Then
                ElseIf piChar < n1215 Then ' 1168..1214
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar = n1217 Or piChar = n1219 Or piChar = n1223 Or piChar = n1227 Then
                    piChar = piChar + One
                ElseIf piChar < n1232 Then
                ElseIf piChar < n1259 Then ' 1232..1258
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar < n1262 Then
                ElseIf piChar < n1269 Then ' 1262..1268
                    If piChar Mod Two = Zero Then piChar = piChar + One
                ElseIf piChar = n1272 Then
                    piChar = n1273
                End If
            End If
        Else ' 1329..1366 | 4256..4293
            piChar = piChar + n48
        End If
    ElseIf piChar < n8048 Then ' 4294..8047
        If piChar < n7680 Then
        ElseIf piChar < n7829 Then ' 7680..7828
            If piChar Mod Two = Zero Then piChar = piChar + One
        ElseIf piChar < n7840 Then
        ElseIf piChar < n7929 Then ' 7840..7928
            If piChar Mod Two = Zero Then piChar = piChar + One
        ElseIf piChar < n7944 Then
        ElseIf piChar < n7952 Then ' 7944..7951
            piChar = piChar - Eight
        ElseIf piChar < n7960 Then
        ElseIf piChar < n7966 Then ' 7960..7965
            piChar = piChar - Eight
        ElseIf piChar < n7976 Then
        ElseIf piChar < n7984 Then ' 7976..7983
            piChar = piChar - Eight
        ElseIf piChar < n7992 Then
        ElseIf piChar < n8000 Then ' 7992..7999
            piChar = piChar - Eight
        ElseIf piChar < n8008 Then
        ElseIf piChar < n8014 Then ' 8008..8013
            piChar = piChar - Eight
        ElseIf piChar < n8025 Then
        ElseIf piChar < n8032 Then ' 8025..8031
            If piChar Mod Two Then piChar = piChar - Eight
        ElseIf piChar < n8040 Then
        ElseIf piChar < n8048 Then ' 8040..8047
            piChar = piChar - Eight
        End If
   ElseIf piChar < n8120 Then '..8119
   Else ' piChar < n9424 Then '..9423
        If piChar = n8120 Or piChar = n8121 Then
            piChar = piChar - Eight
        ElseIf piChar = n8122 Or piChar = n8123 Then
            piChar = piChar - n74
        ElseIf piChar < n8136 Then
        ElseIf piChar < n8140 Then ' 8136..8139
            piChar = piChar - n86
        ElseIf piChar < n8152 Then
        ElseIf piChar = n8152 Or piChar = n8153 Then
            piChar = piChar - Eight
        ElseIf piChar = n8154 Or piChar = n8155 Then
            piChar = piChar - n100
        ElseIf piChar < n8168 Then
        ElseIf piChar = n8168 Or piChar = n8169 Then
            piChar = piChar - Eight
        ElseIf piChar = n8170 Or piChar = n8171 Then
            piChar = piChar - n112
        ElseIf piChar = n8172 Then
            piChar = n8165
        ElseIf piChar < n8184 Then
        ElseIf piChar = n8184 Or piChar = n8185 Then
            piChar = piChar - n128
        ElseIf piChar = n8186 Or piChar = n8187 Then
            piChar = piChar - n126
        ElseIf piChar < n8544 Then
        ElseIf piChar < n8560 Then ' 8544..8559
            piChar = piChar + n16
        ElseIf n9397 < piChar Then 'And piChar < n9424
            piChar = piChar + n26
        End If
    End If
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

'Public Function IsGoodLCaseI() As Boolean
'    Dim i As Long
'    Dim iConv As Long
'    For i = -32768 To 65535
'        iConv = i
'        LukeLCaseI iConv
'        IsGoodLCaseI = AscW(LCase$(ChrW$(i))) = iConv
'        Debug.Assert IsGoodLCaseI
'        If Not IsGoodLCaseI Then Exit Function
'    Next
'End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

' QuickSort History

' Sir Charles Antony Richard Hoare (Tony Hoare or C.A.R. Hoare, born 11th
' January, 1934) is a British computer scientist, probably best known for
' the development of Quicksort (or Hoaresort), the world's most widely used
' sorting algorithm.

' The quicksort algorithm was developed by Hoare in 1960 while working for
' the small British scientific computer manufacturer Elliott Brothers.

' He also developed Hoare logic for verifying program correctness, and the
' formal language Communicating Sequential Processes (CSP) used to specify
' the interactions of concurrent processes (including the Dining philosophers
' problem) and the inspiration for the Occam programming language.
'
' A quicksort sorts by employing a divide and conquer strategy to divide a
' list into two sub-lists.

' The steps are:

' 1. Pick an element, called a pivot, from the list.

' 2. Reorder the list so that all elements which are less than the pivot
'    come before the pivot and so that all elements greater than the pivot
'    come after it (equal values can go either way).

'    After this partitioning, the pivot is in its final position.
'    This is called the partition operation.

' 3. Recursively sort the sub-list of lesser elements and the sub-list of
'    greater elements.

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
