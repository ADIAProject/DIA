Attribute VB_Name = "mApiErrMsg"
Option Explicit

Private Declare Function FormatMessage _
                          Lib "kernel32.dll" _
                              Alias "FormatMessageA" (ByVal dwFlags As Long, _
                                                      lpSource As Any, _
                                                      ByVal dwMessageId As Long, _
                                                      ByVal dwLanguageId As Long, _
                                                      ByVal lpBuffer As String, _
                                                      ByVal nSize As Long, _
                                                      Arguments As Long) As Long

Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000

Public Function ApiErrorText(ByVal errNum As Long) As String

Dim Msg                                 As String
Dim nRet                                As Long

    Msg = String$(1024, vbNullChar)
    nRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, errNum, 0&, Msg, Len(Msg), ByVal 0&)

    If nRet Then
        ApiErrorText = Replace$(Left$(Msg, nRet), vbNewLine, vbNullString)
    Else
        ApiErrorText = "Error (" & errNum & ") not defined."

    End If

End Function
