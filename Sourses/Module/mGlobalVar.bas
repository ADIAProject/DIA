Attribute VB_Name = "mGlobalVar"
Option Explicit

Public strAppPath                 As String
Public strAppPathBackSL           As String

' Переменная для формы показа произвольного сообщения
Public lngShowMessageResult       As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub GetCurAppPath
'! Description (Описание)  :   [процедура получения глобальных переменных путей программы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub GetCurAppPath()
    strAppPath = App.Path
    strAppPathBackSL = BackslashAdd2Path(strAppPath)
End Sub
