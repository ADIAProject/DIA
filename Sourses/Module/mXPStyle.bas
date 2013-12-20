Attribute VB_Name = "mXPStyle"
Option Explicit

' Модуль для инициализации стиля XP+ в программах, требуется файл манифеста в ресурсах программы
Public Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Public mbAeroEnabled As Boolean
Public mbAppThemed   As Boolean

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsAeroEnabled
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsAeroEnabled() As Boolean

    Dim GlassState As Long

    If APIFunctionPresent("DwmIsCompositionEnabled", "dwmapi.dll") Then
        Call DwmIsCompositionEnabled(GlassState)
        IsAeroEnabled = CBool(GlassState)
    End If

End Function
