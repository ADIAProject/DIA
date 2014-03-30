Attribute VB_Name = "mApiErrMsg"
Option Explicit

Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ApiErrorText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   errNum (Long)
'!--------------------------------------------------------------------------------
Public Function ApiErrorText(ByVal errNum As Long) As String

    Dim Msg  As String
    Dim nRet As Long

    Msg = FillNullChar(1024)
    nRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, errNum, 0&, Msg, Len(Msg), ByVal 0&)

    If nRet Then
        ApiErrorText = Replace$(Left$(Msg, nRet), vbNewLine, vbNullString)
    Else
        ApiErrorText = "Error (" & errNum & ") not defined."
    End If

End Function
