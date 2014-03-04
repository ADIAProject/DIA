Attribute VB_Name = "mFontProp"
Option Explicit

'Переменная, хранящая значение кодировки шрифта, считывается из языкового файла, применяется ко всем элементам интерфейса
Public lngFont_Charset       As Long

' Шрифт основной формы и шрифта подсказок
Public strFontMainForm_Name  As String
Public lngFontMainForm_Size  As Long
' Шрифт других форм
Public strFontOtherForm_Name As String
Public lngFontOtherForm_Size As Long

'Переменные, хранящие изменяемые свойства шрифта для кнопок пакетов драйверов
Public lngFontBtn_Color      As Long
Public strFontBtn_Name       As String
Public miFontBtn_Size        As Integer
Public mbFontBtn_Italic      As Boolean
Public mbFontBtn_Underline   As Boolean
Public mbFontBtn_Strikethru  As Boolean
Public mbFontBtn_Bold        As Boolean

'Переменные, хранящие изменяемые свойства шрифта ssTAB1
Public lngFontTab_Color      As Long
Public strFontTab_Name       As String
Public miFontTab_Size        As Integer
Public mbFontTab_Italic      As Boolean
Public mbFontTab_Underline   As Boolean
Public mbFontTab_Strikethru  As Boolean
Public mbFontTab_Bold        As Boolean

'Переменные, хранящие изменяемые свойства шрифта ssTAB2
Public lngFontTab2_Color     As Long
Public strFontTab2_Name      As String
Public miFontTab2_Size       As Integer
Public mbFontTab2_Italic     As Boolean
Public mbFontTab2_Underline  As Boolean
Public mbFontTab2_Strikethru As Boolean
Public mbFontTab2_Bold       As Boolean

'Переменные, хранящие изменяемые свойства шрифта для всплывающих подсказок (ToolTip)
Public lngFontTT_Color       As Long
Public strFontTT_Name        As String
Public miFontTT_Size         As Integer
Public mbFontTT_Italic       As Boolean
Public mbFontTT_Underline    As Boolean
Public mbFontTT_Strikethru   As Boolean
Public mbFontTT_Bold         As Boolean

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetBtnStatusFontProperties
'! Description (Описание)  :   [Установка свойств шрифта для Объекта (кнопка пакета)]
'! Parameters  (Переменные):   ctlObject (Object)
'!--------------------------------------------------------------------------------
Public Sub SetBtnStatusFontProperties(ctlObject As Object)

    With ctlObject
        .ForeColor = lngFontBtn_Color
        .Font.Name = strFontBtn_Name
        .Font.Size = miFontBtn_Size
        .Font.Underline = mbFontBtn_Underline
        .Font.Strikethrough = mbFontBtn_Strikethru
        .Font.Bold = mbFontBtn_Bold
        .Font.Italic = mbFontBtn_Italic
        .Font.Charset = lngFont_Charset
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetBtnFontProperties
'! Description (Описание)  :   [Установка свойств шрифта для Объекта (прочие кнопки)]
'! Parameters  (Переменные):   ctlObject (Object)
'!--------------------------------------------------------------------------------
Public Sub SetBtnFontProperties(ctlObject As Object)

    With ctlObject
        .Font.Name = strFontMainForm_Name
        .Font.Size = lngFontMainForm_Size
        .Font.Charset = lngFont_Charset
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetTTFontProperties
'! Description (Описание)  :   [Установка свойств шрифта для Объекта (ToolTip)]
'! Parameters  (Переменные):   ctlObject (Object)
'!--------------------------------------------------------------------------------
Public Sub SetTTFontProperties(ctlObject As Object)

    With ctlObject
        .ForeColor = lngFontTT_Color
        With .Font
            .Name = strFontTT_Name
            .Size = miFontTT_Size
            .Underline = mbFontTT_Underline
            .Strikethrough = mbFontTT_Strikethru
            .Bold = mbFontTT_Bold
            .Italic = mbFontTT_Italic
            .Charset = lngFont_Charset
        End With
    End With

End Sub


