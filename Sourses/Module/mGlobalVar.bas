Attribute VB_Name = "mGlobalVar"
Option Explicit

Public strAppPath                 As String
Public strAppPathBackSL           As String
Public strAppEXEName              As String

' Переменная для формы показа произвольного сообщения
Public lngShowMessageResult       As Long

'Per the excellent advice of Kroc (camendesign.com), a custom UserMode variable is less prone to errors than the usual
' Ambient.UserMode value supplied to ActiveX controls.  This fixes a problem where ActiveX controls sometimes think they
' are being run in a compiled EXE, when actually their properties are just being written as part of .exe compiling.
Public g_UserModeFix As Boolean

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub GetMyAppValue
'! Description (Описание)  :   [процедура получения глобальных переменных путей программы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub GetMyAppProperties()
    strProductVersion = App.Major & strDot & App.Minor & strDot & App.Revision
    strProductName = App.ProductName & " v." & strProductVersion & " @" & App.CompanyName
    strAppPath = App.Path
    strAppPathBackSL = BackslashAdd2Path(strAppPath)
    strAppEXEName = App.EXEName
End Sub

