Attribute VB_Name = "mChooseFont"
Option Explicit

'Переменные, хранящие изменяемые свойства шрифта
Public strDialog_FontName      As String
Public miDialog_FontSize       As Integer
Public mbDialog_Italic         As Boolean
Public mbDialog_Underline      As Boolean
Public mbDialog_Strikethru     As Boolean
Public mbDialog_Bold           As Boolean
Public lngDialog_Color         As Long
Public lngDialog_Language      As Long
Public lngDialog_Charset       As Long

'Переменные, хранящие изменяемые свойства шрифта TAB
Public strDialogTab_FontName   As String
Public miDialogTab_FontSize    As Integer
Public mbDialogTab_Italic      As Boolean
Public mbDialogTab_Underline   As Boolean
Public mbDialogTab_Strikethru  As Boolean
Public mbDialogTab_Bold        As Boolean
Public lngDialogTab_Language   As Boolean
Public lngDialogTab_Color      As Long

'Переменные, хранящие изменяемые свойства шрифта TAB2
Public strDialogTab2_FontName  As String
Public miDialogTab2_FontSize   As Integer
Public mbDialogTab2_Italic     As Boolean
Public mbDialogTab2_Underline  As Boolean
Public mbDialogTab2_Strikethru As Boolean
Public mbDialogTab2_Bold       As Boolean
Public lngDialogTab2_Color     As Long
Public lngDialogTab2_Language  As Long

' Установка свойст шрифта для Объекта (кнопки)
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetButtonProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   ctlObject (Object)
'!--------------------------------------------------------------------------------
Public Sub SetButtonProperties(ctlObject As Object)

    With ctlObject
        .Font.Name = strDialog_FontName
        .Font.Size = miDialog_FontSize
        .Font.Underline = mbDialog_Underline
        .Font.Strikethrough = mbDialog_Strikethru
        .Font.Bold = mbDialog_Bold
        .Font.Italic = mbDialog_Italic
        .Font.Charset = lngDialog_Charset
    End With

End Sub
