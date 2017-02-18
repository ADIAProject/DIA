VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLickety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'=======
'Lickety
'=======
'Load large text file to String array.  Alternative to using Split().
'
'Handles ANSI or Unicode, various line delimiters.  Last line can be
'terminated by LineDelim or not.
'
'See comments in LoadBlock() below for credits.

Private Const DEFAULT_BLOCK As Long = 65536 'In bytes.
Private Const BIGGER_BLOCK_FACTOR As Long = 50
Private Const USE_BIGGER_BLOCK As Long = DEFAULT_BLOCK * BIGGER_BLOCK_FACTOR

Public Enum BlockSizes
    BlockSizeAuto = 0
    BlockSizeDefault = DEFAULT_BLOCK
    'Or any other positive Long value > Len(LineDelim)... anything smaller
    'is treated as BlockSizeAuto.
End Enum

Private Function LoadBlock( _
    ByVal File As Integer, _
    ByRef Remaining As Long, _
    ByRef LeftOver As String, _
    ByVal BlockSize As Long, _
    ByVal LineDelim As String, _
    ByVal Unicode As Boolean) As String()
Attribute LoadBlock.VB_UserMemId = 1610809345
    
    'Based on SplitB05/SplitC01 by Donald, donald@xbeat.net, 20020601
    'at: http://www.xbeat.net/vbspeed/c_Split.htm
    
    Dim LineDelimLen As Long
    Dim BlockLines() As String
    Dim BlockLen As Long
    Dim PosTokens() As Long
    Dim Block() As Byte
    Dim BlockString As String
    Dim PosToken As Long
    Dim LastToken As Long
    Dim TokenIndex As Long
    Dim TokenLen As Long 'In bytes.

    LineDelimLen = Len(LineDelim)
    If LineDelimLen = 0 Then
        ReDim Block(Remaining - 1)
        Get #File, , Block
        ReDim BlockLines(0)
        BlockLines(0) = StrConv(Block, vbUnicode)
        Remaining = 0
    Else
        If Remaining >= BlockSize Then
            BlockLen = BlockSize
            Remaining = Remaining - BlockSize
        Else
            BlockLen = Remaining
            Remaining = 0
        End If
        ReDim Block(BlockLen - 1)
        Get #File, , Block
        If Unicode Then
            BlockString = Block
        Else
            BlockString = StrConv(Block, vbUnicode)
        End If
        Erase Block
        BlockString = LeftOver & BlockString
        BlockLen = Len(BlockString)
        ReDim PosTokens(BlockLen \ LineDelimLen) 'Max possible tokens ("BlockLines").
        PosTokens(0) = 1
        PosToken = InStr(BlockString, LineDelim)
        Do While PosToken
            LastToken = LastToken + 1
            PosTokens(LastToken) = PosToken + LineDelimLen
            PosToken = InStr(PosToken + LineDelimLen, BlockString, LineDelim)
        Loop
        If LastToken = 0 Then
            BlockLines = Split(vbNullString)
        Else
            ReDim BlockLines(LastToken - 1)
            For TokenIndex = 0 To LastToken - 1
                TokenLen = PosTokens(TokenIndex + 1) - LineDelimLen - PosTokens(TokenIndex)
                BlockLines(TokenIndex) = Mid$(BlockString, PosTokens(TokenIndex), TokenLen)
            Next
        End If
        TokenLen = BlockLen - PosTokens(LastToken) + 1
        LeftOver = Mid$(BlockString, PosTokens(LastToken), TokenLen)
    End If
    LoadBlock = BlockLines
End Function

Public Function SplitFile( _
    ByVal FileName As String, _
    Optional ByVal BlockSize As BlockSizes = BlockSizeAuto, _
    Optional ByVal LineDelim As String = vbNewLine, _
    Optional ByVal Unicode As Boolean = False, _
    Optional ByVal SkipBOM As Boolean = False) As String()
    
    Dim Lines() As String
    Dim F As Integer
    Dim Remaining As Long
    Dim BlockLines() As String
    Dim LeftOver As String
    Dim UBoundLinesOld As Long
    Dim LineIndex As Long
    
    Lines = Split(vbNullString) 'Valid String array with Ubound = -1.
    F = FreeFile(0)
    Open FileName For Binary Access Read As #F
    Remaining = LOF(F)
    If SkipBOM Then
        If Unicode Then
            'Skip over UTF-16LE BOM:
            Remaining = Remaining - 2
            Seek #F, 3
        End If
    End If
    'Process "auto blocksize" - normally 0 but anything < Len(LineDelim):
    If BlockSize <= Len(LineDelim) Then
        If Remaining > USE_BIGGER_BLOCK Then
            BlockSize = Remaining \ BIGGER_BLOCK_FACTOR
        Else
            BlockSize = BlockSizeDefault
        End If
    End If
    Do While Remaining > 0
        BlockLines = LoadBlock(F, Remaining, LeftOver, BlockSize, LineDelim, Unicode)
        UBoundLinesOld = UBound(Lines)
        If UBound(BlockLines) >= 0 Then
            ReDim Preserve Lines(UBoundLinesOld + UBound(BlockLines) + 1)
            For LineIndex = 0 To UBound(BlockLines)
                Lines(UBoundLinesOld + LineIndex + 1) = BlockLines(LineIndex)
            Next
        ElseIf Remaining = 0 And Len(LeftOver) > 0 Then
            'Case where input file did not have a final terminating LineDelim:
            ReDim Preserve Lines(UBoundLinesOld + 1)
            Lines(UBoundLinesOld + 1) = LeftOver
        End If
    Loop
    Close #F
    SplitFile = Lines
End Function

