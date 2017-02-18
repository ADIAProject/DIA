VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCMDArguments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'************************************************************************************************************
'************************************************************************************************************
'************************************************************************************************************
'*******         EliteXP® CommandLine Parser Engine
'*******         This class contains two methods of parsing command lines
'*******         Classical method is generally slow as it contains extensive loops as well as checks
'*******         API method depends upon the CommandLineToArgv API which
'*******                 a) Returns the no of arguments in the string
'*******                 b) Returns the pointer to the pointer of the strings array
'*******         The pointer are cycled using loop and the strings are copied to the variable using
'*******                 a) StrLenPtrW() which gives the length of unicode string pointer
'*******                                 Original API: lstrlenW
'*******                                         Declare Function StrLenPtrW lib "kernel32.dll" alias "lstrlenW" (Byval lpString As Long) As Long
'*******                 b) StrCpyPtrToStrW() which copies the string from ptr to string buffer
'*******                                 Original API [lstrcpyW]
'*******                                 Declare Function StrCpyPtrToStrW lib "kernel32.dll" alias "lstrcpyW"(LpStringDest As String, Byval LpStringSource As Long) As Long
'************************************************************************************************************
'************************************************************************************************************
'******If you want the generalised form of API and hate repeated typing of API,
'                               **********************************************
'******Downlod WinAPI32.tlb at: http://www.elitexp.we.bs/products/winapi32.zip
'                               **********************************************
'a) After downloading the zip file, extract the file at any place
'b) Run your vb project and Select References from Project menu
'c) Browse to the path you extract winapi32.tlb . Click on Ok.
'd) You will get tons of APIS not needed to used in WinAPI loader.
'e) For comments and updates mail at: elitexp2008@gmail.com
'f) Website under construction: http://www.elitexp.we.bs
'************************************************************************************************************
'************************************************************************************************************
'************************************************************************************************************
'************************************************************************************************************
Private cFm_CommandLine As String
Private Declare Function CommandLineToArgvW Lib "shell32.dll" (ByVal nInput As Long, ByRef retCount As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function SysAllocString Lib "oleaut32.dll" (ByVal nString As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property CommandLine
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get CommandLine() As String
    CommandLine = cFm_CommandLine
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property CommandLine
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sNewString (String)
'!--------------------------------------------------------------------------------
Public Property Let CommandLine(ByVal sNewString As String)
    cFm_CommandLine = sNewString
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function Arguments
'! Description (Описание)  :   [Get unicode supporting unicode]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function Arguments() As Collection

    Dim lngRet   As Long
    Dim retCount As Long
    Dim rt       As String
    Dim ii       As Integer
    Dim sPtr     As Long
    Dim cmdLine  As String
    Dim retColl  As Collection

    Set retColl = New Collection
    cmdLine = cFm_CommandLine
    lngRet = CommandLineToArgvW(StrPtr(cmdLine & vbNullChar), retCount)

    If retCount Then

        For ii = 0 To retCount - 1
            CopyMemoryLong VarPtr(sPtr), lngRet + ii * 4, 4
            CopyMemoryLong VarPtr(rt), VarPtr(SysAllocString(sPtr)), 4
            retColl.Add rt
        Next

    End If

    If lngRet <> 0 Then
        LocalFree lngRet
        Set Arguments = retColl
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ArgumentsClassical
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function ArgumentsClassical() As Collection
Attribute ArgumentsClassical.VB_Description = "Returns the collection of commandline Parsed By using loops and classical Methods"

    Dim Cmd         As String
    Dim cmdLen      As Long
    Dim cmdLine     As String
    Dim retColl     As Collection
    Dim Append      As Boolean
    Dim InvertStart As Boolean
    Dim CurPos      As Long
    Dim curChar     As String
    Dim CurArg      As String

    On Error Resume Next
    
    'Initialize the command Line and the count
    Set retColl = New Collection
    cmdLine = cFm_CommandLine
    Cmd = cFm_CommandLine
    cmdLen = Len(Cmd)
    '****Declare Return Collection
    '****Declare Used Variables***
    '***************************************
    Append = True

    For CurPos = 1 To cmdLen
        curChar = Mid$(cmdLine, CurPos, 1)

        Select Case curChar

            Case strSpace, vbTab
                Append = False

                If InvertStart Then
                    Append = True
                End If

            Case strQuotes
                InvertStart = Not InvertStart

                If InvertStart Then
                    Append = True
                End If

            Case Else
                Append = True
        End Select

        '***********Checks if to append to the argument
        Select Case Append

            Case True

                If curChar <> strQuotes Then
                    CurArg = CurArg & curChar
                End If

            Case False
                retColl.Add CurArg
                CurArg = vbNullString
        End Select

        DoEvents
    Next

    If LenB(CurArg) Then
        retColl.Add CurArg
    End If

    Set ArgumentsClassical = retColl

    On Error GoTo 0

End Function
