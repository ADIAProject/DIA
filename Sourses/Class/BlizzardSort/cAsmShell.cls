VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAsmShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  cAsmShell.cls - Accelerated ShellSort Algorithm    -©Rd-

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

' You are free to use any part or all of this code even for
' commercial purposes under the agreement that you receive
' absolutely no warranties expressed or implied.

' These sorting routines have the following features:

'  Safe to use with all compiler optimizations.
'  Capable of sorting arrays of millions of string items.
'  Capable of sorting in ascending and descending order.
'  Capable of case-sensitive and case-insensitive criteria.
'  Capable of any Lbound with equal or greater Ubound.

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

' For comments, suggestions or bug reports you can contact
' me at: rde123•bigpond•com.

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

' This shellsort algorithm is unbelievably optimized and is
' a solid performer on all data with no known weaknesses.

' Turning off strict compatibility with VB's StrComp function
' on case-insensitive operations will improve performance.

Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
Private Declare Sub CopyMemByR Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)

Private Type SAFEARRAY1D       ' SAFEARRAY Header
    cDims      As Integer      ' Count of dimensions in this array
    fFeatures  As Integer      ' Bitfield flags indicating attributes of the array
    cbElements As Long         ' Byte size of each element of the array
    cLocks     As Long         ' Number of times the array has been locked without corresponding unlock. The cLocks field is a reference count that indicates how many times the array has been locked. When there is no lock, you're not supposed to access the array data, which is located in pvData.
    pvData     As Long         ' Pointer to the start of the array data (use only if cLocks > 0)
    cElements  As Long         ' Count of elements in this dimension
    lLbound    As Long         ' The lower-bounding signed index number of this dimension
    lUbound    As Long         ' The upper-bounding signed index number of this dimension
End Type

Private Const VT_BYREF = &H4000&     ' Tests whether the InitedArray routine was passed a Variant that contains an array, rather than directly an array in the former case ptr already points to the SA structure. Thanks to Monte Hansen for this fix
Private Const FADF_FIXEDSIZE = &H10& ' Array may not be resized or reallocated

Public Enum eSortOrder
    Descending = -1&
    DefaultOrder = 0&
    Ascending = 1&
    #If False Then
        Dim Descending, DefaultOrder, Ascending
    #End If
End Enum

Public Enum eCompareMethod
    BinaryCompare = 0&
    TextCompare = 1&
    #If False Then
        Dim BinaryCompare, TextCompare
    #End If
End Enum

Private TwisterBuf() As Long ' Twister copymemory buffer
Private StackLB() As Long    ' Pending boundary stacks
Private StackUB() As Long

' Used for unsigned arithmetic
Private Const DW_MSB = &H80000000 ' DWord Most Significant Bit
Private Const n20K As Long = 20000

' More efficient repeated use of numeric literals
Private Const n0 = 0&, n1 = 1&, n2 = 2&, n4 = 4&, n8 = 8&, n16 = 16&

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Const Default_Order As Long = Ascending
Private Const Default_Method As Long = BinaryCompare

Private mMethod As eCompareMethod
Private mOrder As eSortOrder

Private iaLCase() As Integer

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private aAsmShB() As Byte
Private aAsmShI() As Byte

Private lpAsmShB As Long
Private lpAsmShI As Long

Private lpOldFnc1 As Long
Private lpOldFnc2 As Long

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

'  The following properties should be set before sorting

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:  SortOrder [Public Property]
'  Purpose:    Determines the order in which the strings are sorted
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Property Get SortOrder() As eSortOrder
    If mOrder = DefaultOrder Then mOrder = Default_Order
    SortOrder = mOrder
End Property

Property Let SortOrder(ByVal eNewOrder As eSortOrder)
    If eNewOrder = DefaultOrder Then
        If mOrder = DefaultOrder Then mOrder = Default_Order
    Else
        mOrder = eNewOrder
    End If
End Property

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:  SortMethod [Public Property]
'  Purpose:    Determines the way strings are compared
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Property Get SortMethod() As eCompareMethod
    SortMethod = mMethod
End Property

Property Let SortMethod(ByVal eNewMethod As eCompareMethod)
    mMethod = eNewMethod
End Property

' ¤¤ ShellSort ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

' This shell algorithm is founded on a very solid performer
' that was originally developed by vb2themax, and optimized
' further by several very talented coders, before I twisted
' a little more out of it in this accelerated version.

' This algorithm is unbelievably optimized and is a solid
' performer on all data with no known weaknesses.

' Turning off strict compatibility with VB's StrComp function
' will improve performance on case-insensitive operations.

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Sub sShell(sA() As String, Optional ByVal StrCompCompatibility As Boolean = True)
   If SortMethod = BinaryCompare Then
      Shell_Binary sA, SortOrder
   Else
      Shell_Insens sA, iaLCase, SortOrder
      If StrCompCompatibility Then strTwist4 sA, LBound(sA), UBound(sA)
   End If
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
' Don't move these procedures, add public procedures anywhere,
' or add private procedues above them. See Class_Initialize.
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Sub Shell_Binary(sA() As String, ByVal iComp As Long)
';;;;;;;;;;; Asm code patched on init ;;;;;;;;;;;;
End Sub

Private Sub Shell_Insens(sA() As String, iAlc() As Integer, ByVal iComp As Long)
';;;;;;;;;;; Asm code patched on init ;;;;;;;;;;;;
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

' Important : DON'T SAVE THIS FILE IN ANYPAD! Escaped machine code may get mangled.
' Many text editors will substitute these extended ascii characters breaking the asm (they range from 128 to 255).

Private Sub Class_Initialize() ' Please do not edit the following binary machine code
    Dim lpObjVFT As Long

    AsmEscStr2Bin aAsmShB, "žÕ‹ìƒì”æÁœÖ×ÒÑÓ‹´õŒ‹ž…Û¡„£€€‹Ë¦ƒù‚‚ˆÈ€€‹óŒ‰æõø‹Åƒøè€ü¯‹ž…ÛŒõˆƒÆ„Éõ‹ôëÐ¸€€ä€»ÁôÇ‹Î”†…Òõ‰¶œ†‰" & _
                           "–ƒÆ„Ù‹žÉëçÀë÷ä‹ÁÈ‹ÿ‹Žœ†…Ûõ†ÉìÈõõë‹øºÈ‹”†…ÒõÚ‰œ†‰”¾ŠÏ‹œ¾ÉÈøæ…ëéÈùæ‰Îõøƒù‚‚±€€€‹ÿ³ŸÀ‹ÐÑâÀƒµÂ»Áòõ¼" & _
                           "ëŽ‰ýü³Ò¹ðƒ€€€÷ñ‰®ÅìÁà‚‰Åëô‹õøƒð‹ëÿ‹†‰Åð‹Ùþ«ýô‹—‹íÞ×‹ø³É‚æ‹„¹æ»„°‘õæ…Àô¶…ƒÁ‚ëìßë æ«„‘ßÀæ³Åù•‹›ÿ‰“" & _
                           "‹ß«ýÉô»ýøò‡‹ú—‹ÅðëÃ‹Ûû‹Åð‰‡ƒÑÆ„»õüõ¦‹Åìƒøõ……³ÀÛÙÚßÜÞæ‹åÝÂ„Œ€¸€€€ƒëì"
    
    AsmEscStr2Bin aAsmShI, "žÕ‹ìƒì˜æÁœÖ×ÒÑÓ‹´õ‹ž…Û¡„ß€€‹Ë†ù€€€Â…Ð€€‹´óŒ‰õô‹õÚŒ‹ž…Û„½€€‹Ë“ƒù‚‚¢¤€€‹óŒ‰õ³ø‹Å”ƒø€Üü±‹" & _
                           "ÿ‹ž…™ÛõˆƒÆ„É–õôëÐ¸€È€€»ÁôÇ‹”†…Òõí‰œ†‰–ƒÆ²„‹žÉëçÀïëä‹ÁÈ‹ÿ‹œ†…Ûõ†ØÉÈõõë‹õøÈ‹”†…Ò´õ‰œ†‰”•¾Ï‹œ¾ÉÈ" & _
                           "Ìø…ëéÈùæ‰õøƒù‚ã‚¨€€€‹ÿ¾³À‹ÐÑâÀêƒÂ»ÁòõÖ¼Ž‰ýü³Òá¹ƒ€€€÷ñÝ‰ÅèÁà‚‰ÖÅð‹õøƒð×‹ÿ‹†‰Åì³‹þ«ýð‹—£‹ÞÖ×Ó‹õ×ô‹ø" & _
                           "³Û³ÀŠ³Éæ‹„¹æ‘‹„Ææ‹œ‘€æ»„Þõæ³…Àô…ƒÁ‚£ëäÛßÞë¢€æ«„ÞÛßÞÀæ³Å”ù•‹›ÿ‰“‹ß«ýÉð»ýøò‡‹ú—‹Åìë¯‹Ûû‹Åì‰‡ƒÑÆ„»õüõ’" & _
                           "‹ÅèƒøÝ…íÿÿÿ³ÀÀÛÙÚßÞæË‹åÝÂ€¸°€€€ëì"

   ' VFT > public routines in declared order, then private routines in declared order
    CopyMemByR lpObjVFT, ByVal ObjPtr(Me), n4 ' Must be implemented within a class
    lpAsmShB = lpObjVFT + &H34& '&H1C + 6 * 4 ' First public routine = ObjVFT + &H1C
    lpAsmShI = lpObjVFT + &H38& '&H1C + 7 * 4

    CopyMemByR lpOldFnc1, ByVal lpAsmShB, n4
    CopyMemByR ByVal lpAsmShB, VarPtr(aAsmShB(0)), n4
    CopyMemByR lpOldFnc2, ByVal lpAsmShI, n4
    CopyMemByR ByVal lpAsmShI, VarPtr(aAsmShI(0)), n4

    mOrder = Default_Order
    mMethod = Default_Method
    UnsBuildTextCompTable iaLCase
End Sub

Private Sub Class_Terminate()
    CopyMemByR ByVal lpAsmShB, lpOldFnc1, n4
    CopyMemByR ByVal lpAsmShI, lpOldFnc2, n4
End Sub

' ¤¤ Twister ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
' Procedure: strTwist4 [Public Sub]
'
' Description: Stable Insert/Binary hybrid using CopyMemory
'
' Used internally as a finisher for case-insensitive operations
' to remove differences between TextComp and VB's StrComp.
'
' Available as a public routine for refresh sorting operations.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Sub strTwist4(sA() As String, ByVal lbA As Long, ByVal ubA As Long)
    Dim walk As Long, find As Long, midd As Long
    Dim base As Long, ceil As Long, item As String
    Dim cast As Long, mezz As Long, run As Long
    Dim lpStr As Long, lp As Long, lpB As Long
    Dim idx As Long
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim eComp As eSortOrder
    walk = (ubA - lbA) + n1                      ' Grab array item count
    If walk < n2 Then Exit Sub                   ' If nothing to do then exit
    eComp = SortOrder                            ' Initialize compare variable
    InitializeBuffers n20K                       ' Initialize working buffers
    lpStr = VarPtr(item)                                          ' Cache pointer to the string variable
    lp = Sum(VarPtr(sA(lbA)), -(lbA * n4))                        ' Cache pointer to the array
    lpB = VarPtr(TwisterBuf(n0))                                  ' Cache pointer to the buffer
    walk = lbA: mezz = ubA                                        ' Initialize our walker variables
    Do Until walk = mezz ' ----==============================---- ' Do the twist while there's more items
        walk = walk + n1                                          ' Walk up the array and use binary search to insert each item down into the sorted lower array
        CopyMemByV lpStr, Sum(lp, walk * n4), n4                  ' Grab current value into item
        find = walk                                               ' Default to current position
        ceil = walk - n1                                          ' Set ceiling to current position - 1
        base = lbA                                                ' Set base to lower bound
        Do While StrComp(sA(ceil), item, mMethod) = eComp   '  .  ' While current item must move down
            midd = (base + ceil) \ n2                             ' Find mid point
            Do Until StrComp(sA(midd), item, mMethod) = eComp     ' Step back up if equal or below
                base = midd + n1                                  ' Bring up the base
                midd = (base + ceil) \ n2                         ' Find mid point
                If midd = ceil Then Exit Do                       ' If we're up to ceiling
            Loop                                                  ' Out of loop > target pos
            find = midd                                           ' Set provisional to new ceiling
            If find = base Then Exit Do                           ' If we're down to base
            ceil = midd - n1                                      ' Bring down the ceiling
        Loop '-Twister v4 ©Rd-     .      . ...  .             .  ' Out of stable binary search loops
        If (find < walk) Then                                     ' If current item needs to move down
            CopyMemByV lpStr, Sum(lp, find * n4), n4
            run = walk + n1
            Do Until run > mezz Or run - walk > n20K              ' Runner do loop
                If Not (StrComp(item, sA(run), mMethod) = eComp) Then Exit Do
                run = run + n1
            Loop: cast = (run - walk)
            CopyMemByV lpB, Sum(lp, walk * n4), cast * n4         ' Grab current value(s)
            CopyMemByV Sum(lp, (find + cast) * n4), Sum(lp, find * n4), (walk - find) * n4 ' Move up items
            CopyMemByV Sum(lp, find * n4), lpB, cast * n4         ' Re-assign current value(s) into found pos
            If cast > n1 Then
                If Not run > mezz Then
                    idx = idx + n1
                    StackLB(idx) = run - n1  ' Will increment back
                    StackUB(idx) = mezz
                End If
                walk = find
                mezz = find + cast - n1
        End If: End If
        If walk = mezz Then
            If idx Then
                walk = StackLB(idx)
                mezz = StackUB(idx)
                idx = idx - n1
    End If: End If: Loop           ' Out of walker do loop
    CopyMemByR ByVal lpStr, n0, n4 ' De-reference pointer to item variable
    ' ----=================----
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:   InitializeBuffers [Private Function]
'  Purpose:     Initializes pending runner stacks
'               Initializes the Twisters runner buffer
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Function InitializeBuffers(ByVal StackSize As Long) As Long
    ReDim StackLB(StackSize) As Long    ' Stack to hold pending lower boundries
    ReDim StackUB(StackSize) As Long    ' Stack to hold pending upper boundries
    ReDim TwisterBuf(StackSize) As Long ' This is a cache used when moving ranges
End Function

' ¤¤ Sum ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

' Enables valid addition and subtraction of unsigned long ints.
' Treats lPtr as an unsigned long and returns an unsigned long.
' Allows safe arithmetic operations on memory address pointers.
' Assumes valid pointer and pointer offset.

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Function Sum(ByVal lPtr As Long, ByVal lOffset As Long) As Long
    If lOffset > n0 Then
        If lPtr And DW_MSB Then ' if ptr < 0
           Sum = lPtr + lOffset ' ignors > unsigned int max
        ElseIf (lPtr Or DW_MSB) < -lOffset Then
           Sum = lPtr + lOffset ' result is below signed int max
        Else                    ' result wraps to min signed int
           Sum = (lPtr + DW_MSB) + (lOffset + DW_MSB)
        End If
    ElseIf lOffset = n0 Then
        Sum = lPtr
    Else 'If lOffset < 0 Then
        If (lPtr And DW_MSB) = n0 Then ' if ptr > 0
           Sum = lPtr + lOffset ' ignors unsigned int < zero
        ElseIf (lPtr - DW_MSB) >= -lOffset Then
           Sum = lPtr + lOffset ' result is above signed int min
        Else                    ' result wraps to max signed int
           Sum = (lOffset - DW_MSB) + (lPtr - DW_MSB)
        End If
    End If
End Function

' ¤¤ Inited Array ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

' This function determines if the passed array is initialized,
' and if so will return -1.

' It will also optionally indicate whether the array can be
' redimmed; in which case it will return -2.

' If the array is uninitialized (never redimmed or has been
' erased) it will return 0 (zero).

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Function InitedArray(Arr, lbA As Long, ubA As Long, Optional ByVal bTestRedimable As Boolean) As Long
    ' Thanks to Francesco Balena who solved the Variant headache,
    ' and to Monte Hansen for the ByRef fix
    Dim tSA As SAFEARRAY1D, lpSA As Long
    Dim lOffset As Long, iDataType As Integer
    On Error GoTo UnInit
    CopyMemByR iDataType, Arr, n2                       ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then           ' if a valid array was passed
        CopyMemByR lpSA, ByVal Sum(VarPtr(Arr), n8), n4 ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array
        If (iDataType And VT_BYREF) Then                ' see whether the function was passed a Variant that contains an array, rather than directly an array in the former case lpSA already points to the SA structure. Thanks to Monte Hansen for this fix
            CopyMemByR lpSA, ByVal lpSA, n4             ' lpSA is a discripter (pointer) to the safearray structure
        End If
        InitedArray = (lpSA <> n0)
        If InitedArray Then
            CopyMemByR tSA.cDims, ByVal lpSA, n4
            If bTestRedimable Then ' Return -2 if redimmable
                InitedArray = InitedArray + ((tSA.fFeatures And FADF_FIXEDSIZE) <> FADF_FIXEDSIZE)
            End If '-©Rd-
            lOffset = n16 + ((tSA.cDims - n1) * n8)
            CopyMemByR tSA.cElements, ByVal Sum(lpSA, lOffset), n8
            tSA.lUbound = tSA.lLbound + tSA.cElements - n1
            If (lbA < tSA.lLbound) Then lbA = tSA.lLbound
            If (ubA > tSA.lUbound) Then ubA = tSA.lUbound
    End If: End If
UnInit:
End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Sub AsmEscStr2Bin(b() As Byte, s As String)
    Dim c As Long, r As Long
    Dim i As Long, j As Long
    Dim p As Byte, m As Byte
    j = Len(s): i = (j \ 8): r = j Mod 8
    If r = 0& Then r = i * 7 - 1& Else r = r + i * 7 - 2&
    ReDim b(r) As Byte: m = Asc(Mid$(s, 1&, 1&)): c = 2&
    For i = 0 To r
        b(i) = Asc(Mid$(s, i + c, 1&))
        If (m And (2 ^ p)) = 0 Then b(i) = b(i) And 127
        p = (i + 1) Mod 7
        If p = 0 Then
            c = c + 1&: If i + c > j Then Exit For
            m = Asc(Mid$(s, i + c, 1&))
        End If
    Next i
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  UnsBuildTextCompTable   [Private Sub]
'
'  Purpose:  Initializes the unicode lookup table to enable
'            in-line case-insensitive textual comparison.
'
'            "_", "@", etc, list before number "0"
'            0, ¼, ½, ¾, ¹, 1, ², 2, ³, 3, 4, etc
'            ã , ä, å, a, a, b, b, ç, ç, c, c, etc
'
'  This procedure has been modified to enable unsigned access
'  to the table through a zero offset base pointer to lb.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Sub UnsBuildTextCompTable(ia() As Integer)
   Dim i As Integer, j As Integer
   ReDim ia(-32768 To 32767) As Integer
   For i = -32768 To -256
      ia(i + 32768) = i
   Next
   For i = -255 To -1
      j = i
      TextCompLookup j
      ia(i + 32768) = j   ' Unsigned build
   Next
   For i = 1 To 12539
      j = i
      TextCompLookup j
      ia(i - 32768) = j
   Next
   For i = 12540 To 32766
      ia(i - 32768) = i
   Next
   ia(-1) = i '32767
End Sub

Private Sub TextCompLookup(iChr As Integer)
   If iChr < -255 Then
   ElseIf iChr > 12539 Then
   ElseIf iChr < -161 Then
      If iChr < -210 Then
         If iChr < -233 Then
            Select Case iChr
               Case Is = -255: iChr = 62   ' !
               Case Is = -254: iChr = 75   ' "
               Case Is = -253: iChr = 40   ' #
               Case Is = -252: iChr = 103  ' $
               Case Is = -251: iChr = 84   ' %
               Case Is = -250: iChr = 41   ' &
               Case Is = -249: iChr = 68   ' '
               Case Is = -248: iChr = 44   ' (
               Case Is = -247: iChr = 45   ' )
               Case Is = -246: iChr = 42   ' *
               Case Is = -245: iChr = 85   ' +
               Case Is = -244: iChr = 73   ' ,
               Case Is = -243: iChr = 86   ' -
               Case Is = -242: iChr = 74   ' .
               Case Is = -241: iChr = 50   ' /
               Case Is = -240: iChr = 108  ' 0
               Case Is = -239: iChr = 112  ' 1
               Case Is = -238: iChr = 113  ' 2
               Case Is = -237: iChr = 114  ' 3
               Case Is = -236: iChr = 115  ' 4
               Case Is = -235: iChr = 116  ' 5
               Case Is = -234: iChr = 117  ' 6
            End Select
         Else
            Select Case iChr
               Case Is = -233: iChr = 118  ' 7
               Case Is = -232: iChr = 119  ' 8
               Case Is = -231: iChr = 120  ' 9
               Case Is = -230: iChr = 63   ' :
               Case Is = -229: iChr = 64   ' ;
               Case Is = -228: iChr = 80   ' <
               Case Is = -227: iChr = 88   ' =
               Case Is = -226: iChr = 81   ' >
               Case Is = -224: iChr = 43   ' @
               Case Is = -223: iChr = 122  ' A
               Case Is = -222: iChr = 123  ' B
               Case Is = -221: iChr = 124  ' C
               Case Is = -220: iChr = 125  ' D
               Case Is = -219: iChr = 126  ' E
               Case Is = -218: iChr = 127  ' F
               Case Is = -217: iChr = 128  ' G
               Case Is = -216: iChr = 129  ' H
               Case Is = -215: iChr = 130  ' I
               Case Is = -214: iChr = 131  ' J
               Case Is = -213: iChr = 132  ' K
               Case Is = -212: iChr = 133  ' L
               Case Is = -211: iChr = 134  ' M
            End Select
         End If
      Else 'If iChr < -161 Then
         If iChr < -185 Then
            Select Case iChr
               Case Is = -210: iChr = 135  ' N
               Case Is = -209: iChr = 136  ' O
               Case Is = -208: iChr = 137  ' P
               Case Is = -207: iChr = 138  ' Q
               Case Is = -206: iChr = 139  ' R
               Case Is = -205: iChr = 140  ' S
               Case Is = -204: iChr = 141  ' T
               Case Is = -203: iChr = 142  ' U
               Case Is = -202: iChr = 143  ' V
               Case Is = -201: iChr = 144  ' W
               Case Is = -200: iChr = 145  ' X
               Case Is = -199: iChr = 146  ' Y
               Case Is = -198: iChr = 147  ' Z
               Case Is = -197: iChr = 46   ' [
               Case Is = -196: iChr = 51   ' \
               Case Is = -195: iChr = 47   ' ]
               Case Is = -194: iChr = 79   ' ^
               Case Is = -193: iChr = 52   ' u-score_
               Case Is = -192: iChr = 69   ' `
               Case Is = -191: iChr = 122  ' a
               Case Is = -190: iChr = 123  ' b
               Case Is = -189: iChr = 124  ' c
               Case Is = -188: iChr = 125  ' d
               Case Is = -187: iChr = 126  ' e
               Case Is = -186: iChr = 127  ' f
            End Select
         Else
            Select Case iChr
               Case Is = -185: iChr = 128  ' g
               Case Is = -184: iChr = 129  ' h
               Case Is = -183: iChr = 130  ' i
               Case Is = -182: iChr = 131  ' j
               Case Is = -181: iChr = 132  ' k
               Case Is = -180: iChr = 133  ' l
               Case Is = -179: iChr = 134  ' m
               Case Is = -178: iChr = 135  ' n
               Case Is = -177: iChr = 136  ' o
               Case Is = -176: iChr = 137  ' p
               Case Is = -175: iChr = 138  ' q
               Case Is = -174: iChr = 139  ' r
               Case Is = -173: iChr = 140  ' s
               Case Is = -172: iChr = 141  ' t
               Case Is = -171: iChr = 142  ' u
               Case Is = -170: iChr = 143  ' v
               Case Is = -169: iChr = 144  ' w
               Case Is = -168: iChr = 145  ' x
               Case Is = -167: iChr = 146  ' y
               Case Is = -166: iChr = 147  ' z
               Case Is = -165: iChr = 48   ' {
               Case Is = -164: iChr = 54   ' |
               Case Is = -163: iChr = 49   ' }
               Case Is = -162: iChr = 67   ' ~
            End Select
         End If
      End If
   ElseIf iChr < 32 Then
   ElseIf iChr < 256 Then
      If iChr < 60 Then
         If iChr < 45 Then
            Select Case iChr
               Case Is = 32: iChr = 38    ' space
               Case Is = 33: iChr = 62    ' !
               Case Is = 34: iChr = 75    ' "
               Case Is = 35: iChr = 40    ' #
               Case Is = 36: iChr = 103   ' $
               Case Is = 37: iChr = 84    ' %
               Case Is = 38: iChr = 41    ' &
               Case Is = 39: iChr = 68    ' '
               Case Is = 40: iChr = 44    ' (
               Case Is = 41: iChr = 45    ' )
               Case Is = 42: iChr = 42    ' *
               Case Is = 43: iChr = 85    ' +
               Case Is = 44: iChr = 73    ' ,
            End Select
         Else
            Select Case iChr
               Case Is = 45: iChr = 86    ' -
               Case Is = 46: iChr = 74    ' .
               Case Is = 47: iChr = 50    ' /
               Case Is = 48: iChr = 108   ' 0
               Case Is = 49: iChr = 112   ' 1
               Case Is = 50: iChr = 113   ' 2
               Case Is = 51: iChr = 114   ' 3
               Case Is = 52: iChr = 115   ' 4
               Case Is = 53: iChr = 116   ' 5
               Case Is = 54: iChr = 117   ' 6
               Case Is = 55: iChr = 118   ' 7
               Case Is = 56: iChr = 119   ' 8
               Case Is = 57: iChr = 120   ' 9
               Case Is = 58: iChr = 63    ' :
               Case Is = 59: iChr = 64    ' ;
            End Select
         End If
      ElseIf iChr < 90 Then
         If iChr < 75 Then
            Select Case iChr
               Case Is = 60: iChr = 80    ' <
               Case Is = 61: iChr = 88    ' =
               Case Is = 62: iChr = 81    ' >
               Case Is = 63: iChr = 66    ' ?
               Case Is = 64: iChr = 43    ' @
               Case Is = 65: iChr = 122   ' A
               Case Is = 66: iChr = 123   ' B
               Case Is = 67: iChr = 124   ' C
               Case Is = 68: iChr = 125   ' D
               Case Is = 69: iChr = 126   ' E
               Case Is = 70: iChr = 127   ' F
               Case Is = 71: iChr = 128   ' G
               Case Is = 72: iChr = 129   ' H
               Case Is = 73: iChr = 130   ' I
               Case Is = 74: iChr = 131   ' J
            End Select
         Else
            Select Case iChr
               Case Is = 75: iChr = 132   ' K
               Case Is = 76: iChr = 133   ' L
               Case Is = 77: iChr = 134   ' M
               Case Is = 78: iChr = 135   ' N
               Case Is = 79: iChr = 136   ' O
               Case Is = 80: iChr = 137   ' P
               Case Is = 81: iChr = 138   ' Q
               Case Is = 82: iChr = 139   ' R
               Case Is = 83: iChr = 140   ' S
               Case Is = 84: iChr = 141   ' T
               Case Is = 85: iChr = 142   ' U
               Case Is = 86: iChr = 143   ' V
               Case Is = 87: iChr = 144   ' W
               Case Is = 88: iChr = 145   ' X
               Case Is = 89: iChr = 146   ' Y
            End Select
         End If
      ElseIf iChr < 120 Then
         If iChr < 105 Then
            Select Case iChr
               Case Is = 90: iChr = 147   ' Z
               Case Is = 91: iChr = 46    ' [
               Case Is = 92: iChr = 51    ' \
               Case Is = 93: iChr = 47    ' ]
               Case Is = 94: iChr = 79    ' ^
               Case Is = 95: iChr = 52    ' u-score_
               Case Is = 96: iChr = 69    ' `
               Case Is = 97: iChr = 122   ' a
               Case Is = 98: iChr = 123   ' b
               Case Is = 99: iChr = 124   ' c
               Case Is = 100: iChr = 125  ' d
               Case Is = 101: iChr = 126  ' e
               Case Is = 102: iChr = 127  ' f
               Case Is = 103: iChr = 128  ' g
               Case Is = 104: iChr = 129  ' h
            End Select
         Else
            Select Case iChr
               Case Is = 105: iChr = 130  ' i
               Case Is = 106: iChr = 131  ' j
               Case Is = 107: iChr = 132  ' k
               Case Is = 108: iChr = 133  ' l
               Case Is = 109: iChr = 134  ' m
               Case Is = 110: iChr = 135  ' n
               Case Is = 111: iChr = 136  ' o
               Case Is = 112: iChr = 137  ' p
               Case Is = 113: iChr = 138  ' q
               Case Is = 114: iChr = 139  ' r
               Case Is = 115: iChr = 140  ' s
               Case Is = 116: iChr = 141  ' t
               Case Is = 117: iChr = 142  ' u
               Case Is = 118: iChr = 143  ' v
               Case Is = 119: iChr = 144  ' w
            End Select
         End If
      ElseIf iChr < 150 Then
         If iChr < 135 Then
            Select Case iChr
               Case Is = 120: iChr = 145  ' x
               Case Is = 121: iChr = 146  ' y
               Case Is = 122: iChr = 147  ' z
               Case Is = 123: iChr = 48   ' {
               Case Is = 124: iChr = 54   ' |
               Case Is = 125: iChr = 49   ' }
               Case Is = 126: iChr = 67   ' ~
               Case Is = 127: iChr = 32   ' 
               Case Is = 128: iChr = 98   ' €
               Case Is = 129: iChr = 33   ' 
               Case Is = 130: iChr = 73   ' ‚
               Case Is = 131: iChr = 127  ' ƒ
               Case Is = 132: iChr = 78   ' „
               Case Is = 133: iChr = 59   ' …
               Case Is = 134: iChr = 55   ' †
            End Select
         Else
            Select Case iChr
               Case Is = 135: iChr = 56   ' ‡
               Case Is = 136: iChr = 79   ' ˆ
               Case Is = 137: iChr = 57   ' ‰
               Case Is = 138: iChr = 140  ' Š
               Case Is = 139: iChr = 80   ' ‹
               Case Is = 140: iChr = 136  ' Œ
               Case Is = 141: iChr = 34   ' 
               Case Is = 142: iChr = 147  ' Ž
               Case Is = 143: iChr = 35   ' 
               Case Is = 144: iChr = 36   ' 
               Case Is = 145: iChr = 68   ' '
               Case Is = 146: iChr = 68   ' '
               Case Is = 147: iChr = 75   ' "
               Case Is = 148: iChr = 75   ' "
               Case Is = 149: iChr = 89   ' •
            End Select
         End If
      ElseIf iChr < 182 Then
         If iChr < 166 Then
            Select Case iChr
               Case Is = 150: iChr = 87   ' –
               Case Is = 151: iChr = 87   ' —
               Case Is = 152: iChr = 67   ' ˜
               Case Is = 153: iChr = 105  ' ™
               Case Is = 154: iChr = 140  ' š
               Case Is = 155: iChr = 81   ' ›
               Case Is = 156: iChr = 136  ' œ
               Case Is = 157: iChr = 37   ' 
               Case Is = 158: iChr = 147  ' ž
               Case Is = 159: iChr = 146  ' Ÿ
               Case Is = 160: iChr = 39   ' nobk space
               Case Is = 161: iChr = 61   ' ¡
               Case Is = 162: iChr = 99   ' ¢
               Case Is = 163: iChr = 100  ' £
               Case Is = 164: iChr = 101  ' ¤
               Case Is = 165: iChr = 102  ' ¥
            End Select
         Else
            Select Case iChr
               Case Is = 166: iChr = 58   ' ¦
               Case Is = 167: iChr = 104  ' §
               Case Is = 168: iChr = 60   ' ¨
               Case Is = 169: iChr = 106  ' ©
               Case Is = 170: iChr = 122  ' ª
               Case Is = 171: iChr = 82   ' «
               Case Is = 172: iChr = 95   ' ¬
               Case Is = 173: iChr = 86   ' -
               Case Is = 174: iChr = 107  ' ®
               Case Is = 175: iChr = 53   ' ¯
               Case Is = 176: iChr = 90   ' °
               Case Is = 177: iChr = 91   ' ±
               Case Is = 178: iChr = 113  ' ²
               Case Is = 179: iChr = 114  ' ³
               Case Is = 180: iChr = 70   ' ´
               Case Is = 181: iChr = 96   ' µ
            End Select
         End If
      ElseIf iChr < 215 Then
         If iChr < 198 Then
            Select Case iChr
               Case Is = 182: iChr = 97   ' ¶
               Case Is = 183: iChr = 92   ' ·
               Case Is = 184: iChr = 65   ' ¸
               Case Is = 185: iChr = 112  ' ¹
               Case Is = 186: iChr = 136  ' º
               Case Is = 187: iChr = 83   ' »
               Case Is = 188: iChr = 109  ' ¼
               Case Is = 189: iChr = 110  ' ½
               Case Is = 190: iChr = 111  ' ¾
               Case Is = 191: iChr = 65   ' ¿
               Case Is = 192: iChr = 122  ' À
               Case Is = 193: iChr = 122  ' Á
               Case Is = 194: iChr = 122  ' Â
               Case Is = 195: iChr = 122  ' Ã
               Case Is = 196: iChr = 122  ' Ä
               Case Is = 197: iChr = 122  ' Å
            End Select
         Else
            Select Case iChr
               Case Is = 198: iChr = 121  ' Æ
               Case Is = 199: iChr = 124  ' Ç
               Case Is = 200: iChr = 126  ' È
               Case Is = 201: iChr = 126  ' É
               Case Is = 202: iChr = 126  ' Ê
               Case Is = 203: iChr = 126  ' Ë
               Case Is = 204: iChr = 130  ' Ì
               Case Is = 205: iChr = 130  ' Í
               Case Is = 206: iChr = 130  ' Î
               Case Is = 207: iChr = 130  ' Ï
               Case Is = 208: iChr = 125  ' Ð
               Case Is = 209: iChr = 135  ' Ñ
               Case Is = 210: iChr = 136  ' Ò
               Case Is = 211: iChr = 136  ' Ó
               Case Is = 212: iChr = 136  ' Ô
               Case Is = 213: iChr = 136  ' Õ
               Case Is = 214: iChr = 136  ' Ö
            End Select
         End If
      Else 'If iChr < 256 Then
         If iChr < 235 Then
            Select Case iChr
               Case Is = 215: iChr = 93   ' ×
               Case Is = 216: iChr = 136  ' Ø
               Case Is = 217: iChr = 142  ' Ù
               Case Is = 218: iChr = 142  ' Ú
               Case Is = 219: iChr = 142  ' Û
               Case Is = 220: iChr = 142  ' Ü
               Case Is = 221: iChr = 146  ' Ý
               Case Is = 222: iChr = 141  ' Þ
               Case Is = 223: iChr = 140  ' ß
               Case Is = 224: iChr = 122  ' à
               Case Is = 225: iChr = 122  ' á
               Case Is = 226: iChr = 122  ' â
               Case Is = 227: iChr = 122  ' ã
               Case Is = 228: iChr = 122  ' ä
               Case Is = 229: iChr = 122  ' å
               Case Is = 230: iChr = 121  ' æ
               Case Is = 231: iChr = 124  ' ç
               Case Is = 232: iChr = 126  ' è
               Case Is = 233: iChr = 126  ' é
               Case Is = 234: iChr = 126  ' ê
            End Select
         Else
            Select Case iChr
               Case Is = 235: iChr = 126  ' ë
               Case Is = 236: iChr = 130  ' ì
               Case Is = 237: iChr = 130  ' í
               Case Is = 238: iChr = 130  ' î
               Case Is = 239: iChr = 130  ' ï
               Case Is = 240: iChr = 125  ' ð
               Case Is = 241: iChr = 135  ' ñ
               Case Is = 242: iChr = 136  ' ò
               Case Is = 243: iChr = 136  ' ó
               Case Is = 244: iChr = 136  ' ô
               Case Is = 245: iChr = 136  ' õ
               Case Is = 246: iChr = 136  ' ö
               Case Is = 247: iChr = 94   ' ÷
               Case Is = 248: iChr = 136  ' ø
               Case Is = 249: iChr = 142  ' ù
               Case Is = 250: iChr = 142  ' ú
               Case Is = 251: iChr = 142  ' û
               Case Is = 252: iChr = 142  ' ü
               Case Is = 253: iChr = 146  ' ý
               Case Is = 254: iChr = 141  ' þ
               Case Is = 255: iChr = 146  ' ÿ
            End Select
         End If
      End If
   ElseIf iChr < 497 Then
      If iChr < 284 Then
         If iChr < 270 Then
            Select Case iChr
               Case Is = 256: iChr = 122  ' A
               Case Is = 257: iChr = 122  ' a
               Case Is = 258: iChr = 122  ' A
               Case Is = 259: iChr = 122  ' a
               Case Is = 260: iChr = 122  ' A
               Case Is = 261: iChr = 122  ' a
               Case Is = 262: iChr = 124  ' C
               Case Is = 263: iChr = 124  ' c
               Case Is = 264: iChr = 124  ' C
               Case Is = 265: iChr = 124  ' c
               Case Is = 266: iChr = 124  ' C
               Case Is = 267: iChr = 124  ' c
               Case Is = 268: iChr = 124  ' C
               Case Is = 269: iChr = 124  ' c
            End Select
         Else
            Select Case iChr
               Case Is = 270: iChr = 125  ' D
               Case Is = 271: iChr = 125  ' d
               Case Is = 272: iChr = 125  ' Ð
               Case Is = 273: iChr = 125  ' d
               Case Is = 274: iChr = 126  ' E
               Case Is = 275: iChr = 126  ' e
               Case Is = 276: iChr = 126  ' E
               Case Is = 277: iChr = 126  ' e
               Case Is = 278: iChr = 126  ' E
               Case Is = 279: iChr = 126  ' e
               Case Is = 280: iChr = 126  ' E
               Case Is = 281: iChr = 126  ' e
               Case Is = 282: iChr = 126  ' E
               Case Is = 283: iChr = 126  ' e
            End Select
         End If
      ElseIf iChr < 323 Then
         If iChr < 300 Then
            Select Case iChr
               Case Is = 284: iChr = 128  ' G
               Case Is = 285: iChr = 128  ' g
               Case Is = 286: iChr = 128  ' G
               Case Is = 287: iChr = 128  ' g
               Case Is = 288: iChr = 128  ' G
               Case Is = 289: iChr = 128  ' g
               Case Is = 290: iChr = 128  ' G
               Case Is = 291: iChr = 128  ' g
               Case Is = 292: iChr = 129  ' H
               Case Is = 293: iChr = 129  ' h
               Case Is = 294: iChr = 129  ' H
               Case Is = 295: iChr = 129  ' h
               Case Is = 296: iChr = 130  ' I
               Case Is = 297: iChr = 130  ' i
               Case Is = 298: iChr = 130  ' I
               Case Is = 299: iChr = 130  ' i
            End Select
         Else
            Select Case iChr
               Case Is = 300: iChr = 130  ' I
               Case Is = 301: iChr = 130  ' i
               Case Is = 302: iChr = 130  ' I
               Case Is = 303: iChr = 130  ' i
               Case Is = 304: iChr = 130  ' I
               Case Is = 305: iChr = 130  ' i
               Case Is = 308: iChr = 131  ' J
               Case Is = 309: iChr = 131  ' j
               Case Is = 310: iChr = 132  ' K
               Case Is = 311: iChr = 132  ' k
               Case Is = 313: iChr = 133  ' L
               Case Is = 314: iChr = 133  ' l
               Case Is = 315: iChr = 133  ' L
               Case Is = 316: iChr = 133  ' l
               Case Is = 317: iChr = 133  ' L
               Case Is = 318: iChr = 133  ' l
               Case Is = 321: iChr = 133  ' L
               Case Is = 322: iChr = 133  ' l
            End Select
         End If
      ElseIf iChr < 360 Then
         If iChr < 342 Then
            Select Case iChr
               Case Is = 323: iChr = 135  ' N
               Case Is = 324: iChr = 135  ' n
               Case Is = 325: iChr = 135  ' N
               Case Is = 326: iChr = 135  ' n
               Case Is = 327: iChr = 135  ' N
               Case Is = 328: iChr = 135  ' n
               Case Is = 332: iChr = 136  ' O
               Case Is = 333: iChr = 136  ' o
               Case Is = 334: iChr = 136  ' O
               Case Is = 335: iChr = 136  ' o
               Case Is = 336: iChr = 136  ' O
               Case Is = 337: iChr = 136  ' o
               Case Is = 338: iChr = 136  ' Œ
               Case Is = 339: iChr = 136  ' œ
               Case Is = 340: iChr = 139  ' R
               Case Is = 341: iChr = 139  ' r
            End Select
         Else
            Select Case iChr
               Case Is = 342: iChr = 139  ' R
               Case Is = 343: iChr = 139  ' r
               Case Is = 344: iChr = 139  ' R
               Case Is = 345: iChr = 139  ' r
               Case Is = 346: iChr = 140  ' S
               Case Is = 347: iChr = 140  ' s
               Case Is = 348: iChr = 140  ' S
               Case Is = 349: iChr = 140  ' s
               Case Is = 350: iChr = 140  ' S
               Case Is = 351: iChr = 140  ' s
               Case Is = 352: iChr = 140  ' Š
               Case Is = 353: iChr = 140  ' š
               Case Is = 354: iChr = 141  ' T
               Case Is = 355: iChr = 141  ' t
               Case Is = 356: iChr = 141  ' T
               Case Is = 357: iChr = 141  ' t
               Case Is = 358: iChr = 141  ' T
               Case Is = 359: iChr = 141  ' t
            End Select
         End If
      ElseIf iChr < 418 Then
         If iChr < 376 Then
            Select Case iChr
               Case Is = 360: iChr = 142  ' U
               Case Is = 361: iChr = 142  ' u
               Case Is = 362: iChr = 142  ' U
               Case Is = 363: iChr = 142  ' u
               Case Is = 364: iChr = 142  ' U
               Case Is = 365: iChr = 142  ' u
               Case Is = 366: iChr = 142  ' U
               Case Is = 367: iChr = 142  ' u
               Case Is = 368: iChr = 142  ' U
               Case Is = 369: iChr = 142  ' u
               Case Is = 370: iChr = 142  ' U
               Case Is = 371: iChr = 142  ' u
               Case Is = 372: iChr = 144  ' W
               Case Is = 373: iChr = 144  ' w
               Case Is = 374: iChr = 146  ' Y
               Case Is = 375: iChr = 146  ' y
            End Select
         Else
            Select Case iChr
               Case Is = 376: iChr = 146  ' Ÿ
               Case Is = 377: iChr = 147  ' Z
               Case Is = 378: iChr = 147  ' z
               Case Is = 379: iChr = 147  ' Z
               Case Is = 380: iChr = 147  ' z
               Case Is = 381: iChr = 147  ' Ž
               Case Is = 382: iChr = 147  ' ž
               Case Is = 384: iChr = 123  ' b
               Case Is = 393: iChr = 125  ' Ð
               Case Is = 401: iChr = 127  ' ƒ
               Case Is = 402: iChr = 127  ' ƒ
               Case Is = 407: iChr = 130  ' I
               Case Is = 410: iChr = 133  ' l
               Case Is = 415: iChr = 136  ' O
               Case Is = 416: iChr = 136  ' O
               Case Is = 417: iChr = 136  ' o
            End Select
         End If
      ElseIf iChr < 427 Then
      Else 'If iChr < 497 Then
         If iChr < 472 Then
            Select Case iChr
               Case Is = 427: iChr = 141  ' t
               Case Is = 430: iChr = 141  ' T
               Case Is = 431: iChr = 142  ' U
               Case Is = 432: iChr = 142  ' u
               Case Is = 438: iChr = 147  ' z
               Case Is = 448: iChr = 54   ' |
               Case Is = 451: iChr = 62   ' !
               Case Is = 461: iChr = 122  ' A
               Case Is = 462: iChr = 122  ' a
               Case Is = 463: iChr = 130  ' I
               Case Is = 464: iChr = 130  ' i
               Case Is = 465: iChr = 136  ' O
               Case Is = 466: iChr = 136  ' o
               Case Is = 467: iChr = 142  ' U
               Case Is = 468: iChr = 142  ' u
               Case Is = 469: iChr = 142  ' U
               Case Is = 470: iChr = 142  ' u
               Case Is = 471: iChr = 142  ' U
            End Select
         Else
            Select Case iChr
               Case Is = 472: iChr = 142  ' u
               Case Is = 473: iChr = 142  ' U
               Case Is = 474: iChr = 142  ' u
               Case Is = 475: iChr = 142  ' U
               Case Is = 476: iChr = 142  ' u
               Case Is = 478: iChr = 122  ' A
               Case Is = 479: iChr = 122  ' a
               Case Is = 484: iChr = 128  ' G
               Case Is = 485: iChr = 128  ' g
               Case Is = 486: iChr = 128  ' G
               Case Is = 487: iChr = 128  ' g
               Case Is = 488: iChr = 132  ' K
               Case Is = 489: iChr = 132  ' k
               Case Is = 490: iChr = 136  ' O
               Case Is = 491: iChr = 136  ' o
               Case Is = 492: iChr = 136  ' O
               Case Is = 493: iChr = 136  ' o
               Case Is = 496: iChr = 131  ' j
            End Select
         End If
      End If
   ElseIf iChr < 697 Then
       If iChr = 609 Then iChr = 128  ' g
   ElseIf iChr < 967 Then
      If iChr < 776 Then
         Select Case iChr
            Case Is = 697: iChr = 68   ' '
            Case Is = 698: iChr = 75   ' "
            Case Is = 700: iChr = 68   ' '
            Case Is = 708: iChr = 79   ' ^
            Case Is = 710: iChr = 79   ' ˆ
            Case Is = 712: iChr = 68   ' '
            Case Is = 713: iChr = 53   ' ¯
            Case Is = 714: iChr = 70   ' ´
            Case Is = 715: iChr = 69   ' `
            Case Is = 717: iChr = 52   ' u-score_
            Case Is = 730: iChr = 90   ' °
            Case Is = 732: iChr = 67   ' ˜
            Case Is = 768: iChr = 69   ' `
            Case Is = 769: iChr = 70   ' ´
            Case Is = 770: iChr = 79   ' ^
            Case Is = 771: iChr = 67   ' ~
            Case Is = 772: iChr = 53   ' ¯
            Case Is = 773: iChr = 53   ' ¯
            End Select
         Else
            Select Case iChr
            Case Is = 776: iChr = 60   ' ¨
            Case Is = 778: iChr = 90   ' °
            Case Is = 782: iChr = 75   ' "
            Case Is = 807: iChr = 65   ' ¸
            Case Is = 817: iChr = 52   ' u-score_
            Case Is = 818: iChr = 52   ' u-score_
            Case Is = 894: iChr = 64   ' ;
            Case Is = 915: iChr = 128  ' G
            Case Is = 920: iChr = 141  ' T
            Case Is = 931: iChr = 140  ' S
            Case Is = 934: iChr = 127  ' F
            Case Is = 937: iChr = 136  ' O
            Case Is = 945: iChr = 122  ' a
            Case Is = 946: iChr = 140  ' ß
            Case Is = 948: iChr = 125  ' d
            Case Is = 949: iChr = 126  ' e
            Case Is = 956: iChr = 96   ' µ
            Case Is = 960: iChr = 137  ' p
            Case Is = 963: iChr = 140  ' s
            Case Is = 964: iChr = 141  ' t
            Case Is = 966: iChr = 127  ' f
         End Select
      End If
   ElseIf iChr < 1211 Then
   ElseIf iChr < 1643 Then
      Select Case iChr
         Case Is = 1211: iChr = 129  ' h
         Case Is = 1417: iChr = 63   ' :
         Case Is = 1642: iChr = 84   ' %
      End Select
   ElseIf iChr < 8192 Then
   ElseIf iChr < 10073 Then
      If iChr < 8325 Then
         If iChr < 8226 Then
            Select Case iChr
               Case Is = 8192: iChr = 38   '
               Case Is = 8193: iChr = 38   '
               Case Is = 8194: iChr = 38   '
               Case Is = 8195: iChr = 38   '
               Case Is = 8196: iChr = 38   '
               Case Is = 8197: iChr = 38   '
               Case Is = 8198: iChr = 38   '
               Case Is = 8208: iChr = 86   ' -
               Case Is = 8209: iChr = 86   ' -
               Case Is = 8211: iChr = 87   ' –
               Case Is = 8212: iChr = 87   ' —
               Case Is = 8215: iChr = 88   ' =
               Case Is = 8216: iChr = 71   ' ‘
               Case Is = 8217: iChr = 72   ' ’
               Case Is = 8218: iChr = 73   ' ‚
               Case Is = 8220: iChr = 76   ' “
               Case Is = 8221: iChr = 77   ' ”
               Case Is = 8222: iChr = 78   ' „
               Case Is = 8224: iChr = 55   ' †
               Case Is = 8225: iChr = 56   ' ‡
            End Select
         Else
            Select Case iChr
               Case Is = 8226: iChr = 89   ' •
               Case Is = 8228: iChr = 92   ' ·
               Case Is = 8230: iChr = 59   ' …
               Case Is = 8240: iChr = 57   ' ‰
               Case Is = 8242: iChr = 68   ' '
               Case Is = 8245: iChr = 69   ' `
               Case Is = 8249: iChr = 80   ' ‹
               Case Is = 8250: iChr = 81   ' ›
               Case Is = 8260: iChr = 50   ' /
               Case Is = 8304: iChr = 90   ' °
               Case Is = 8308: iChr = 115  ' 4
               Case Is = 8309: iChr = 116  ' 5
               Case Is = 8310: iChr = 117  ' 6
               Case Is = 8311: iChr = 118  ' 7
               Case Is = 8312: iChr = 119  ' 8
               Case Is = 8319: iChr = 135  ' n
               Case Is = 8320: iChr = 108  ' 0
               Case Is = 8321: iChr = 112  ' 1
               Case Is = 8322: iChr = 113  ' 2
               Case Is = 8323: iChr = 114  ' 3
               Case Is = 8324: iChr = 115  ' 4
            End Select
         End If
      ElseIf iChr < 8501 Then
         If iChr < 8467 Then
            Select Case iChr
               Case Is = 8325: iChr = 116  ' 5
               Case Is = 8326: iChr = 117  ' 6
               Case Is = 8327: iChr = 118  ' 7
               Case Is = 8328: iChr = 119  ' 8
               Case Is = 8329: iChr = 120  ' 9
               Case Is = 8353: iChr = 99   ' ¢
               Case Is = 8356: iChr = 100  ' £
               Case Is = 8359: iChr = 137  ' P
               Case Is = 8364: iChr = 98   ' €
               Case Is = 8450: iChr = 124  ' C
               Case Is = 8455: iChr = 126  ' E
               Case Is = 8458: iChr = 128  ' g
               Case Is = 8459: iChr = 129  ' H
               Case Is = 8460: iChr = 129  ' H
               Case Is = 8461: iChr = 129  ' H
               Case Is = 8462: iChr = 129  ' h
               Case Is = 8464: iChr = 130  ' I
               Case Is = 8465: iChr = 130  ' I
               Case Is = 8466: iChr = 133  ' L
            End Select
         Else
            Select Case iChr
               Case Is = 8467: iChr = 133  ' l
               Case Is = 8469: iChr = 135  ' N
               Case Is = 8472: iChr = 137  ' P
               Case Is = 8473: iChr = 137  ' P
               Case Is = 8474: iChr = 138  ' Q
               Case Is = 8475: iChr = 139  ' R
               Case Is = 8476: iChr = 139  ' R
               Case Is = 8477: iChr = 139  ' R
               Case Is = 8482: iChr = 105  ' ™
               Case Is = 8484: iChr = 147  ' Z
               Case Is = 8488: iChr = 147  ' Z
               Case Is = 8490: iChr = 132  ' K
               Case Is = 8491: iChr = 122  ' Å
               Case Is = 8492: iChr = 123  ' B
               Case Is = 8493: iChr = 124  ' C
               Case Is = 8494: iChr = 126  ' e
               Case Is = 8495: iChr = 126  ' e
               Case Is = 8496: iChr = 126  ' E
               Case Is = 8497: iChr = 127  ' F
               Case Is = 8499: iChr = 134  ' M
               Case Is = 8500: iChr = 136  ' o
            End Select
         End If
      ElseIf iChr < 8709 Then
      ElseIf iChr < 9003 Then
         If iChr < 8776 Then
            Select Case iChr
               Case Is = 8709: iChr = 136  ' Ø
               Case Is = 8722: iChr = 86   ' -
               Case Is = 8723: iChr = 91   ' ±
               Case Is = 8725: iChr = 50   ' /
               Case Is = 8726: iChr = 51   ' \
               Case Is = 8727: iChr = 42   ' *
               Case Is = 8728: iChr = 90   ' °
               Case Is = 8729: iChr = 92   ' ·
               Case Is = 8730: iChr = 143  ' v
               Case Is = 8734: iChr = 119  ' 8
               Case Is = 8739: iChr = 54   ' |
               Case Is = 8745: iChr = 135  ' n
               Case Is = 8758: iChr = 63   ' :
               Case Is = 8764: iChr = 67   ' ~
            End Select
         Else
            Select Case iChr
               Case Is = 8776: iChr = 67   ' ˜
               Case Is = 8801: iChr = 88   ' =
               Case Is = 8804: iChr = 88   ' =
               Case Is = 8805: iChr = 88   ' =
               Case Is = 8810: iChr = 82   ' «
               Case Is = 8811: iChr = 83   ' »
               Case Is = 8901: iChr = 92   ' ·
               Case Is = 8962: iChr = 58   ' ¦
               Case Is = 8963: iChr = 79   ' ^
               Case Is = 8976: iChr = 95   ' ¬
               Case Is = 8992: iChr = 44   ' (
               Case Is = 8993: iChr = 45   ' )
               Case Is = 9001: iChr = 80   ' <
               Case Is = 9002: iChr = 81   ' >
            End Select
         End If
      ElseIf iChr < 9472 Then
      ElseIf iChr < 9789 Then
         If iChr < 9566 Then
            Select Case iChr
               Case Is = 9472: iChr = 86   ' -
               Case Is = 9474: iChr = 58   ' ¦
               Case Is = 9484: iChr = 85   ' +
               Case Is = 9488: iChr = 85   ' +
               Case Is = 9492: iChr = 85   ' +
               Case Is = 9496: iChr = 85   ' +
               Case Is = 9500: iChr = 85   ' +
               Case Is = 9508: iChr = 58   ' ¦
               Case Is = 9516: iChr = 86   ' -
               Case Is = 9524: iChr = 86   ' -
               Case Is = 9532: iChr = 85   ' +
               Case Is = 9552: iChr = 86   ' -
               Case Is = 9553: iChr = 58   ' ¦
               Case Is = 9554: iChr = 85   ' +
               Case Is = 9555: iChr = 85   ' +
               Case Is = 9556: iChr = 85   ' +
               Case Is = 9557: iChr = 85   ' +
               Case Is = 9558: iChr = 85   ' +
               Case Is = 9559: iChr = 85   ' +
               Case Is = 9560: iChr = 85   ' +
               Case Is = 9561: iChr = 85   ' +
               Case Is = 9562: iChr = 85   ' +
               Case Is = 9563: iChr = 85   ' +
               Case Is = 9564: iChr = 85   ' +
               Case Is = 9565: iChr = 85   ' +
            End Select
         Else
            Select Case iChr
               Case Is = 9566: iChr = 58   ' ¦
               Case Is = 9567: iChr = 58   ' ¦
               Case Is = 9568: iChr = 58   ' ¦
               Case Is = 9569: iChr = 58   ' ¦
               Case Is = 9570: iChr = 58   ' ¦
               Case Is = 9571: iChr = 58   ' ¦
               Case Is = 9572: iChr = 86   ' -
               Case Is = 9573: iChr = 86   ' -
               Case Is = 9574: iChr = 86   ' -
               Case Is = 9575: iChr = 86   ' -
               Case Is = 9576: iChr = 86   ' -
               Case Is = 9577: iChr = 86   ' -
               Case Is = 9578: iChr = 85   ' +
               Case Is = 9579: iChr = 85   ' +
               Case Is = 9580: iChr = 85   ' +
               Case Is = 9600: iChr = 53   ' ¯
               Case Is = 9604: iChr = 52   ' u-score_
               Case Is = 9608: iChr = 58   ' ¦
               Case Is = 9612: iChr = 58   ' ¦
               Case Is = 9616: iChr = 58   ' ¦
               Case Is = 9617: iChr = 58   ' ¦
               Case Is = 9618: iChr = 58   ' ¦
               Case Is = 9619: iChr = 58   ' ¦
               Case Is = 9632: iChr = 58   ' ¦
               Case Is = 9788: iChr = 101  ' ¤
            End Select
         End If
     ElseIf iChr = 10072 Then iChr = 54 ' |
     End If
   ElseIf iChr < 12288 Then
   Else 'If iChr < 12540 Then
      Select Case iChr
         Case Is = 12288: iChr = 38  '
         Case Is = 12296: iChr = 80  ' <
         Case Is = 12297: iChr = 81  ' >
         Case Is = 12298: iChr = 82  ' «
         Case Is = 12299: iChr = 83  ' »
         Case Is = 12314: iChr = 46  ' [
         Case Is = 12315: iChr = 47  ' ]
         Case Is = 12539: iChr = 92  ' ·
      End Select
   End If
End Sub

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

' Rd - crYptic but cRaZy!

