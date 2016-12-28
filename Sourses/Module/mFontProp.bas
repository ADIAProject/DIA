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
        With .Font
            .Name = strFontBtn_Name
            .Size = miFontBtn_Size
            .Underline = mbFontBtn_Underline
            .Strikethrough = mbFontBtn_Strikethru
            .Bold = mbFontBtn_Bold
            .Italic = mbFontBtn_Italic
            .Charset = lngFont_Charset
        End With
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetTabProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   ctlObject (Object)
'!--------------------------------------------------------------------------------
Public Sub SetTabProperties(ctlObject As Object)

    With ctlObject
        .ForeColor = lngFontTab_Color
        With .Font
            .Name = strFontTab_Name
            .Size = miFontTab_Size
            .Underline = mbFontTab_Underline
            .Strikethrough = mbFontTab_Strikethru
            .Bold = mbFontTab_Bold
            .Italic = mbFontTab_Italic
            .Charset = lngFont_Charset
        End With
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetTab2Properties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   ctlObject (Object)
'!--------------------------------------------------------------------------------
Public Sub SetTab2Properties(ctlObject As Object)

    With ctlObject
        .ForeColor = lngFontTab2_Color
        With .Font
            .Name = strFontTab2_Name
            .Size = miFontTab2_Size
            .Underline = mbFontTab2_Underline
            .Strikethrough = mbFontTab2_Strikethru
            .Bold = mbFontTab2_Bold
            .Italic = mbFontTab2_Italic
            .Charset = lngFont_Charset
        End With
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetBtnFontProperties
'! Description (Описание)  :   [Установка свойств шрифта для Объекта (кнопка)]
'! Parameters  (Переменные):   ctlObject (Object)
'!--------------------------------------------------------------------------------
Public Sub SetBtnFontProperties(ctlObject As Object)

    With ctlObject
        .ForeColor = lngFontBtn_Color
        With .Font
            .Name = strFontBtn_Name
            .Size = miFontBtn_Size
            .Underline = mbFontBtn_Underline
            .Strikethrough = mbFontBtn_Strikethru
            .Bold = mbFontBtn_Bold
            .Italic = mbFontBtn_Italic
            .Charset = lngFont_Charset
        End With
    End With

End Sub
