Attribute VB_Name = "mChooseFont"
Option Explicit

'Переменные, хранящие изменяемые свойства шрифта
Public strDialog_FontName               As String
Public miDialog_FontSize                As Integer
Public mbDialog_Italic                  As Boolean
Public mbDialog_Underline               As Boolean
Public mbDialog_Strikethru              As Boolean
Public mbDialog_Bold                    As Boolean
Public lngDialog_Color                  As Long
Public lngDialog_Language               As Long
Public lngDialog_Charset                As Long

'Переменные, хранящие изменяемые свойства шрифта TAB
Public strDialogTab_FontName            As String
Public miDialogTab_FontSize             As Integer
Public mbDialogTab_Italic               As Boolean
Public mbDialogTab_Underline            As Boolean
Public mbDialogTab_Strikethru           As Boolean
Public mbDialogTab_Bold                 As Boolean
Public lngDialogTab_Language            As Boolean
Public lngDialogTab_Color               As Long

'Переменные, хранящие изменяемые свойства шрифта TAB2
Public strDialogTab2_FontName           As String
Public miDialogTab2_FontSize            As Integer
Public mbDialogTab2_Italic              As Boolean
Public mbDialogTab2_Underline           As Boolean
Public mbDialogTab2_Strikethru          As Boolean
Public mbDialogTab2_Bold                As Boolean
Public lngDialogTab2_Color              As Long
Public lngDialogTab2_Language           As Long

Public Sub GetButtonProperties(ByVal ButtonName As ctlXpButton)

    With ButtonName
        'Сохранение визуально заданых свойств шрифтов в переменных
        strDialog_FontName = .Font.Name
        miDialog_FontSize = .Font.Size
        mbDialog_Underline = .Font.Underline
        mbDialog_Strikethru = .Font.Strikethrough
        mbDialog_Bold = .Font.Bold
        mbDialog_Italic = .Font.Italic
        lngDialog_Color = .TextColor
        lngDialog_Language = .Font.Charset
    End With

End Sub

Public Sub GetButtonPropertiesJC(ByVal ButtonName As ctlJCbutton)

    With ButtonName
        'Сохранение визуально заданых свойств шрифтов в переменных
        strDialog_FontName = .Font.Name
        miDialog_FontSize = .Font.Size
        mbDialog_Underline = .Font.Underline
        mbDialog_Strikethru = .Font.Strikethrough
        mbDialog_Bold = .Font.Bold
        mbDialog_Italic = .Font.Italic
        lngDialog_Color = .ForeColor
    End With

End Sub

Public Sub GetTabProperties(ButtonName As ctlXpButton)

    With ButtonName
        'Сохранение визуально заданых свойств шрифтов в переменных
        strDialogTab_FontName = .Font.Name
        miDialogTab_FontSize = .Font.Size
        mbDialogTab_Underline = .Font.Underline
        mbDialogTab_Strikethru = .Font.Strikethrough
        mbDialogTab_Bold = .Font.Bold
        mbDialogTab_Italic = .Font.Italic
        lngDialogTab_Color = .TextColor

    End With

End Sub

Public Sub SetButtonProperties(Optional ByVal ButtonName As ctlXpButton, _
                               Optional ByVal ButtonNameJC As ctlJCbutton, _
                               Optional ByVal IsJCButton As Boolean = False)

Dim ctlObject                           As Object

    'Сохранение визуально заданых свойств шрифтов в переменных
    If IsJCButton Then
        Set ctlObject = ButtonNameJC
    Else
        Set ctlObject = ButtonName
    End If

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
